import os

from PyPDF2 import PdfReader

folder_input = "./Input/"
folder_output = "./Output/"

files = [f for f in os.listdir(folder_input) if '.pdf' in f]

pass_dict = dict()
with open(f"{folder_input}pass.txt", 'r') as pass_file:
    for line in pass_file:
        key, value = line.split(":")
        pass_dict[key] = value.replace("\n", "")

result = []

for file in files:
    filename,_ = file.split('.pdf')
    _, pass_key, _ = file.split('_', maxsplit=2)
    reader = PdfReader(folder_input+file)
    reader.decrypt(pass_dict[pass_key])

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    with open(f"{folder_output}{filename}.txt", "w", encoding="utf-8") as f:
        f.write(text)