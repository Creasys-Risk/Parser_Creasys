import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import shutil
from pathlib import Path
import threading

class ParserApp:
    def __init__(self, ventana):
        self.ventana = ventana
        self.ventana.title("Parser Creasys")
        self.ventana.configure(bg='#f0f0f0')
        
        self.scripts = {
            'BTG': 'btg.py',
            'Banchile': 'banchile_inversiones.py',
            'Security': 'security.py',
            'Santander': 'santander.py',
            'Bice': 'bice.py',
            'JP Morgan': 'jp_morgan.py',
        }

        self.carpeta_raiz = None
        self.proceso_activo = False
        self.base_dir = Path(__file__).parent

        self.configurar_estilos()
        self.configurar_ui()

    def configurar_estilos(self):
        estilo = ttk.Style()
        estilo.theme_use('clam')
        estilo.configure('TButton', padding=6, relief='flat', background='#3e76B2', foreground='white')
        estilo.map('TButton', background=[('active', '#3e76B2')])
        estilo.configure('Hover.TButton', padding=6, relief='flat', background='#2e66a2', foreground='white')
        estilo.map('Hover.TButton', background=[('active', '#2e66a2')])
        estilo.configure('TFrame', background='#f0f0f0')
        estilo.configure('TLabel', background='#f0f0f0', foreground='#333333')
        estilo.configure('Red.TButton', background='#a02725')

    def configurar_ui(self):
        marco_principal = ttk.Frame(self.ventana)
        marco_principal.pack(expand=True, fill='both', padx=20, pady=20)

        ttk.Label(marco_principal, text="Seleccione la carpeta raíz con los PDF", font=('Arial', 10, 'bold')).pack(pady=10)

        self.btn_carpeta_raiz = ttk.Button(
            marco_principal,
            text="Buscar Carpeta Raíz",
            command=self.seleccionar_carpeta_raiz
        )
        self.btn_carpeta_raiz.pack(pady=5)

        self.lbl_ruta = ttk.Label(marco_principal, text="", wraplength=400)
        self.lbl_ruta.pack(pady=5)

        self.marco_botones = ttk.Frame(marco_principal)
        self.marco_botones.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(self.marco_botones, bg='#f0f0f0', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.marco_botones, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)

        self.instituciones_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.instituciones_frame, anchor='nw')

        self.instituciones_frame.bind("<Configure>", self.ajustar_scroll)
        self.canvas.bind("<Configure>", self.ajustar_ancho)

        self.progressbar = ttk.Progressbar(marco_principal, mode='indeterminate', length=500)
        self.lbl_proceso = ttk.Label(marco_principal, text="", font=('Arial', 10, 'italic'))
        self.lbl_proceso.pack(pady=5)

        self.boton_salir = ttk.Button(
            marco_principal,
            text="Salir",
            style='Red.TButton',
            command=self.ventana.destroy
        )
        self.boton_salir.pack(side='bottom', pady=15)

    def ajustar_scroll(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        if event.height > self.canvas.winfo_height():
            self.scrollbar.pack(side='right', fill='y')
        else:
            self.scrollbar.pack_forget()

    def ajustar_ancho(self, event):
        self.canvas.itemconfig(1, width=event.width)

    def seleccionar_carpeta_raiz(self):
        self.carpeta_raiz = filedialog.askdirectory()
        if self.carpeta_raiz:
            self.lbl_ruta.config(text=f"Carpeta seleccionada: {self.carpeta_raiz}")
            self.crear_botones_instituciones()
    
    def crear_botones_instituciones(self):
        for widget in self.instituciones_frame.winfo_children():
            widget.destroy()

        instituciones_presentes = []
        user_root = Path(self.carpeta_raiz)
        for entry in user_root.iterdir():
            if entry.is_dir() and entry.name in self.scripts:
                instituciones_presentes.append(entry.name)

        for institucion in instituciones_presentes:
            script_path = self.base_dir / institucion / self.scripts[institucion]
            script_existe = script_path.exists()
            
            user_institucion_dir = user_root / institucion
            pdfs = list(user_institucion_dir.glob('*.pdf'))
            tiene_pdfs = len(pdfs) > 0
            
            estado = 'normal' if script_existe and tiene_pdfs else 'disabled'
            
            btn = ttk.Button(
                self.instituciones_frame,
                text=f"Iniciar Proceso: {institucion}",
                command=lambda inst=institucion: self.iniciar_proceso(inst),
                state=estado,
                style='TButton'
            )
            btn.pack(fill='x', pady=3, padx=20)
            
            btn.bind("<Enter>", lambda e, b=btn: b.configure(style="Hover.TButton"))
            btn.bind("<Leave>", lambda e, b=btn: b.configure(style="TButton"))

    def iniciar_proceso(self, institucion):
        if not self.proceso_activo:
            self.proceso_activo = True
            self.lbl_proceso.config(text=f"Procesando: {institucion}...")
            self.progressbar.pack(pady=15)
            self.progressbar.start()
            threading.Thread(target=lambda: self.procesar_institucion(institucion)).start()
    
    def procesar_institucion(self, institucion):
        try:
            script_dir = self.base_dir / institucion
            user_dir = Path(self.carpeta_raiz) / institucion
            input_dir = script_dir / 'input'
            output_dir = script_dir / 'output'

            input_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)

            pdfs = list(user_dir.glob('*.pdf'))
            txts = list(user_dir.glob('*.txt'))
            archivos_a_procesar = pdfs + txts

            if not pdfs:
                raise Exception("No se encontraron archivos PDF para procesar")
                
            for archivo in archivos_a_procesar:
                shutil.copy2(str(archivo), str(input_dir / archivo.name))

            proceso = subprocess.run(
                ['python', str(script_dir / self.scripts[institucion])],
                cwd=str(script_dir),
                capture_output=True,
                text=True,
                check=True
            )

            excel_files = list(output_dir.glob('*.xlsx'))
            if not excel_files:
                raise Exception("No se generó archivo Excel\nDetalles:\n" + proceso.stderr)
            elif len(excel_files) > 1:
                raise Exception("Múltiples archivos Excel generados:\n" + "\n".join([f.name for f in excel_files]))
            
            informe_generado = excel_files[0]
            shutil.move(str(informe_generado), str(user_dir / informe_generado.name))

            self.ventana.after(0, self.mostrar_resultado, True, f"Proceso completado: {institucion}")

        except subprocess.CalledProcessError as e:
            error_msg = f"Error en ejecución (Código {e.returncode}):\n{e.stderr}"
            self.ventana.after(0, self.mostrar_resultado, False, error_msg)
        except Exception as e:
            self.ventana.after(0, self.mostrar_resultado, False, str(e))
        finally:
            try:
                for f in input_dir.glob('*'):
                    if f.is_file():
                        f.unlink(missing_ok=True)

                for subdir in output_dir.rglob('*'):
                    if subdir.is_dir() and subdir.name.lower() == 'textos':
                        for txt_file in subdir.glob('*.txt'):
                            txt_file.unlink(missing_ok=True)
                
                for txt_file in output_dir.rglob('*.txt'):
                    txt_file.unlink(missing_ok=True)
                    
            except Exception as e:
                print(f"Error en limpieza: {e}")
            
            self.ventana.after(0, self.finalizar_proceso)

    def mostrar_resultado(self, exitoso, mensaje):
        if exitoso:
            messagebox.showinfo("Proceso Finalizado", mensaje)
        else:
            messagebox.showerror("Error en Proceso", mensaje)

    def finalizar_proceso(self):
        self.progressbar.stop()
        self.progressbar.pack_forget()
        self.lbl_proceso.config(text="")
        self.proceso_activo = False
        self.crear_botones_instituciones()

if __name__ == "__main__":
    ventana = tk.Tk()
    app = ParserApp(ventana)
    ventana.mainloop()