import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import threading
from BTG.btg import BTG_Parser
from BanchileInversiones.banchile_inversiones import BanChile_Parser
from Security.security import Security_Parser
from Santander.santander import Santander_Parser
from Bice.bice import Bice_parser
from JP_Morgan.jp_morgan import JPM_Parser

class ParserApp:
    def __init__(self, ventana: tk.Tk):
        self.ventana = ventana
        self.ventana.title("Parser Creasys")
        self.ventana.configure(bg='#f0f0f0')
        self.scripts = {
            'BTG': BTG_Parser,
            'BanchileInversiones': BanChile_Parser,
            'Security': Security_Parser,
            'Santander': Santander_Parser,
            'Bice': Bice_parser,
            'JP Morgan': JPM_Parser,
        }
        self.carpeta_raiz = None
        self.proceso_activo = False
        self.base_dir = Path(__file__).parent
        self.selecciones = {}
        self.todos_seleccionados = False

        self.configurar_estilos()
        self.configurar_ui()

    def configurar_estilos(self):
        estilo = ttk.Style()
        estilo.theme_use('clam')
        estilo.configure('TButton', padding=6, relief='flat', background='#3e76B2', foreground='white')
        estilo.map('TButton', background=[('active', '#3e76B2')])
        estilo.configure('Red.TButton', background='#a02725')
        estilo.configure('TFrame', background='#f0f0f0')
        estilo.configure('TLabel', background='#f0f0f0', foreground='#333333')
        estilo.configure('Treeview', rowheight=25)
        estilo.configure('Treeview.Heading', font=('Arial', 9, 'bold'))

    def configurar_ui(self):
        marco_principal = ttk.Frame(self.ventana)
        marco_principal.pack(expand=True, fill='both', padx=20, pady=20)
        ttk.Label(marco_principal, text="Seleccione la carpeta raíz con los PDF", font=('Arial', 10, 'bold')).pack(pady=10)
        self.btn_carpeta_raiz = ttk.Button(marco_principal, text="Buscar Carpeta Raíz", command=self.seleccionar_carpeta_raiz)
        self.btn_carpeta_raiz.pack(pady=5)
        self.lbl_ruta = ttk.Label(marco_principal, text="", wraplength=400)
        self.lbl_ruta.pack(pady=5)
        self.crear_tabla_instituciones(marco_principal)
        self.progressbar = ttk.Progressbar(marco_principal, mode='indeterminate', length=500)
        self.lbl_proceso = ttk.Label(marco_principal, text="", font=('Arial', 10, 'italic'))
        self.lbl_proceso.pack(pady=5)
        self.boton_procesar = ttk.Button(marco_principal, text="Procesar seleccionados", command=self.procesar_seleccionados, state='disabled')
        self.boton_procesar.pack(pady=10)
        self.boton_salir = ttk.Button(marco_principal, text="Salir", style='Red.TButton', command=self.ventana.destroy)
        self.boton_salir.pack(side='bottom', pady=15)

    def crear_tabla_instituciones(self, parent):
        contenedor = ttk.Frame(parent)
        contenedor.pack(fill='both', expand=True, pady=10)
        self.tree = ttk.Treeview(contenedor, columns=('check', 'institucion', 'estado'), show='headings', selectmode='none')
        self.tree.column('check', width=50, anchor='center', stretch=False)
        self.tree.column('institucion', width=150, anchor='w', stretch=False)
        self.tree.column('estado', width=300, anchor='w')
        self.tree.heading('check', text='☐', command=self.toggle_todos)
        self.tree.heading('institucion', text='Institución')
        self.tree.heading('estado', text='Estado')
        scrollbar_vertical = ttk.Scrollbar(contenedor, orient="vertical", command=self.tree.yview)
        scrollbar_horizontal = ttk.Scrollbar(contenedor, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)
        scrollbar_vertical.pack(side='right', fill='y')
        scrollbar_horizontal.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)
        self.tree.tag_configure('error', foreground='red')
        self.tree.tag_configure('advertencia', foreground='orange')
        self.tree.tag_configure('procesado', foreground='green')
        self.tree.bind('<Button-1>', self.toggle_checkbox)

    def toggle_todos(self):
        self.todos_seleccionados = not self.todos_seleccionados
        new_value = '☑' if self.todos_seleccionados else '☐'
        self.tree.heading('check', text=new_value)
        for item in self.tree.get_children():
            if self.tree.item(item, 'tags') != ('disabled',):
                self.selecciones[item] = self.todos_seleccionados
                self.tree.set(item, 'check', '☑' if self.todos_seleccionados else '☐')
        self.boton_procesar['state'] = 'normal' if self.todos_seleccionados else 'disabled'

    def toggle_checkbox(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == 'cell':
            columna = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            if columna == '#1' and self.tree.item(item, 'tags') != ('disabled',):
                current_value = self.selecciones.get(item, False)
                new_value = not current_value
                self.selecciones[item] = new_value
                self.tree.set(item, 'check', '☑' if new_value else '☐')
                any_selected = any(self.selecciones.values())
                self.boton_procesar['state'] = 'normal' if any_selected else 'disabled'

    def actualizar_estado_institucion(self, institucion, mensaje, tipo=None):
        for child in self.tree.get_children():
            if self.tree.item(child)['values'][1] == institucion:
                self.tree.set(child, 'estado', mensaje)
                if tipo: self.tree.item(child, tags=(tipo,))
                break

    def seleccionar_carpeta_raiz(self):
        self.carpeta_raiz = filedialog.askdirectory()
        if self.carpeta_raiz:
            self.lbl_ruta.config(text=f"Carpeta seleccionada: {self.carpeta_raiz}")
            self.actualizar_lista_instituciones()

    def actualizar_lista_instituciones(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.selecciones.clear()
        instituciones_presentes = [entry.name for entry in Path(self.carpeta_raiz).iterdir() 
                                if entry.is_dir() and entry.name in self.scripts]
        for institucion in instituciones_presentes:
            user_dir = Path(self.carpeta_raiz) / institucion
            pdfs = list(user_dir.glob('*.pdf'))
            estado = 'Listo para procesar' if pdfs else 'Sin archivos PDF'
            tags = ('disabled', 'advertencia') if not pdfs else ()
            check_estado = '☑' if pdfs else '⛔'
            item = self.tree.insert('', 'end', values=(check_estado, institucion, estado), tags=tags)
            self.selecciones[item] = bool(pdfs)
        any_selected = any(self.selecciones.values())
        self.boton_procesar['state'] = 'normal' if any_selected else 'disabled'

    def procesar_seleccionados(self):
        if not self.proceso_activo:
            self.proceso_activo = True
            self.progressbar.pack(pady=15)
            self.progressbar.start()
            instituciones = [self.tree.item(item)['values'][1] 
                           for item in self.tree.get_children() if self.selecciones.get(item, False)]
            threading.Thread(target=self.procesar_instituciones, args=(instituciones,)).start()

    def procesar_instituciones(self, instituciones):
        try:
            for institucion in instituciones:
                self.ventana.after(0, self.lbl_proceso.config, {'text': f"Procesando: {institucion}..."})
                self.ventana.after(0, self.actualizar_estado_institucion, institucion, "Procesando...")
                try:
                    user_dir: Path = Path(self.carpeta_raiz) / institucion
                    pdfs = list(user_dir.glob('*.pdf'))
                    if not pdfs:
                        raise Exception("No hay archivos PDF para procesar")
                    
                    parser_func = self.scripts.get(institucion)
                    if not parser_func:
                        raise Exception(f"Parser no disponible para {institucion}")
                    
                    parser_func(input=user_dir, output=user_dir)
                    
                    excels = list(user_dir.glob('*.xlsx'))
                    if not excels:
                        raise Exception("No se generó archivo de salida")
                    
                    self.ventana.after(0, self.actualizar_estado_institucion, institucion, "Procesado correctamente", 'procesado')
                except Exception as e:
                    self.ventana.after(0, self.actualizar_estado_institucion, institucion, str(e)[:150], 'error')
            self.ventana.after(0, self.finalizar_proceso)
        except Exception as e:
            self.ventana.after(0, messagebox.showerror, "Error", str(e))
            self.ventana.after(0, self.finalizar_proceso)

    def finalizar_proceso(self):
        self.progressbar.stop()
        self.progressbar.pack_forget()
        self.lbl_proceso.config(text="")
        self.proceso_activo = False

if __name__ == "__main__":
    ventana = tk.Tk()
    app = ParserApp(ventana)
    ventana.mainloop()