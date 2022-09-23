"""Переводим файлы из pdf в jpg"""
from pdf2image import convert_from_path
import os


path=r"C:\Users\ulito\Desktop\training\patterns\ex_d\schemes\\"

file_in_dir = os.listdir(path)

def pdftojpg(name, pat):
    pages = convert_from_path(f"{pat}{name}", 200, poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
    for page in pages:
        page.save(f'{pat}{name.split(".")[0]}.jpg', 'JPEG')

for file in file_in_dir:
    if file.endswith(".pdf"):
        pdftojpg(file,path)