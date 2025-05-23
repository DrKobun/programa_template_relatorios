import pandas as pd
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from tkinter import Tk, Label, Entry, Button
from docx import Document
from tkinter import filedialog
import os

caminho_imagem = ""

def update_label_imagem():
    caminho_imagem.split()
    label_imagem.config(text=f"Nome da imagem: {caminho_imagem.split()[-1]}")

def add_imagem():
    global caminho_imagem
    caminho_imagem = os.path.abspath(filedialog.askopenfilename())
    print(f"Imagem escolhida com sucesso! ", caminho_imagem)
    update_label_imagem()

def generate_document():
    # caminho_imagem = filedialog.askopenfilename()
    
    nome_projeto = entry_nome_projeto.get()
    processo = entry_processo.get()
    uf = entry_uf.get()
    tipo_convenente = entry_tipo_convenente.get()
    titulo_objeto = entry_titulo_objeto.get()
    nome_documento = entry_nome_documento.get() + ".docx"

    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    # caminho_imagem = 'C:\\Users\\walyson.ferreira\\Downloads\\LOGO-MIDR-removebg-preview.png'
    
    doc = DocxTemplate(caminho_template)
    
    # Update the placeholder code
    width = float(entry_width.get())  # Get the width from the user input
    height = float(entry_height.get())  # Get the height from the user input
    imagem = InlineImage(doc, caminho_imagem, width=Mm(width), height=Mm(height))  # Set width and height in millimeters

    context = {
        'nome_projeto': nome_projeto,
        'processo': processo,
        'uf': uf,
        'tipo_convenente': tipo_convenente,
        'titulo_objeto': titulo_objeto,
        'imagem': imagem
    }
    
    doc.render(context)
    doc.save(nome_documento)
    root.destroy()

# Create the tkinter window
root = Tk()
root.title("Preencher template")

# Create labels and entry fields
Label(root, text="Nome do Projeto:").grid(row=0, column=0, padx=10, pady=5)
entry_nome_projeto = Entry(root, width=40)
entry_nome_projeto.grid(row=0, column=1, padx=10, pady=5)

Label(root, text="Processo:").grid(row=1, column=0, padx=10, pady=5)
entry_processo = Entry(root, width=40)
entry_processo.grid(row=1, column=1, padx=10, pady=5)

Label(root, text="UF:").grid(row=2, column=0, padx=10, pady=5)
entry_uf = Entry(root, width=40)
entry_uf.grid(row=2, column=1, padx=10, pady=5)

Label(root, text="Tipo do Convenente:").grid(row=3, column=0, padx=10, pady=5)
entry_tipo_convenente = Entry(root, width=40)
entry_tipo_convenente.grid(row=3, column=1, padx=10, pady=5)

Label(root, text="Título do Objeto:").grid(row=4, column=0, padx=10, pady=5)
entry_titulo_objeto = Entry(root, width=40)
entry_titulo_objeto.grid(row=4, column=1, padx=10, pady=5)

Label(root, text="Nome do documento:").grid(row=5, column=0, padx=10, pady=5)
entry_nome_documento = Entry(root, width=40)
entry_nome_documento.grid(row=5, column=1, padx=10, pady=5)

# Add labels and entry fields for width and height side by side in the same column
frame_dimensions = Label(root)
frame_dimensions.grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky="w")

Label(frame_dimensions, text="Width (mm):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_width = Entry(frame_dimensions, width=10)
entry_width.grid(row=0, column=1, padx=5, pady=5, sticky="w")

Label(frame_dimensions, text="Height (mm):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
entry_height = Entry(frame_dimensions, width=10)
entry_height.grid(row=0, column=3, padx=5, pady=5, sticky="w")

# Create a button to generate the document
Button(root, text="Gerar arquivo", command=generate_document).grid(row=7, column=0, pady=10)
Button(root, text="Adicionar Imagem", command=add_imagem).grid(row=7, column=1, pady=10)

label_imagem = Label(root, text=caminho_imagem)
label_imagem.grid(row=8, column=0, pady=10, columnspan=2)

# main loop do tkinter
root.mainloop()