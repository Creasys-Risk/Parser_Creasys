from datetime import date
import os
import re

from PyPDF2 import PdfReader
import pandas as pd

folder_input = "./Input/"
folder_output = "./Output/"

files = [f for f in os.listdir(folder_input) if '.pdf' in f]

info_cartera: list[dict] = []
info_movimientos: list[dict] = []

for file in files:
    filename,_ = file.split('.pdf')
    reader = PdfReader(folder_input+file)

    text = ''

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text += page.extract_text()

    with open(f"{folder_output}{filename}.txt", "w", encoding="utf-8") as f:
        f.write(text)
