from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import win32com.client
import subprocess

wdFormatPDF = 17

def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if caminho_arquivo:
        label.config(text=f"Arquivo selecionado: {caminho_arquivo}")
        return caminho_arquivo
    else:
        label.config(text="Nenhum arquivo selecionado")
        return None

def converter_pdf_para_word():
    caminho_arquivo = selecionar_arquivo()
    if not caminho_arquivo:
        return
    
    # Verificar se a extensão do arquivo é .pdf
    extensao = os.path.splitext(caminho_arquivo)[1]
    if extensao.lower() != ".pdf":
        messagebox.showerror("Erro", "Por favor, selecione um arquivo PDF.")
        return
    
    # Determinar o diretório onde o script Python está localizado
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    
    # Criar uma pasta específica dentro do diretório atual
    pasta_destino = os.path.join(diretorio_atual, "docs")
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    # Nome do arquivo de saída
    nome_arquivo_saida = os.path.splitext(os.path.basename(caminho_arquivo))[0] + ".docx"
    
    # Caminho completo para o arquivo de saída dentro da pasta especificada
    caminho_saida = os.path.join(pasta_destino, nome_arquivo_saida)
    
    # Converter PDF para DOCX
    try:
        cv = Converter(caminho_arquivo)
        cv.convert(caminho_saida, start=0, end=None)  # Adicionado start=0 e end=None
        cv.close()
        messagebox.showinfo("Sucesso", f"Arquivo convertido e salvo em: {caminho_saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def converter_word_para_pdf():
    inputFile = filedialog.askopenfilename(filetypes=[("WORD Files", "*.docx"), ("WORD Files 97-2003", "*.doc")])

    if not inputFile:
        return

    # Nome do arquivo sem extensão
    file_name = os.path.splitext(os.path.basename(inputFile))[0]

    # Determinar o diretório onde o script Python está localizado
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    
    # Criar uma pasta específica dentro do diretório atual
    pasta_destino = os.path.join(diretorio_atual, "docs")
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    # Caminho completo para o arquivo de saída dentro da pasta especificada
    outputFile = os.path.abspath(os.path.join(pasta_destino, f"{file_name}.pdf"))

    # Verifica se o arquivo de saída já existe e o remove se necessário
    if os.path.exists(outputFile):
        os.remove(outputFile)

    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(inputFile)
        doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso e salvo em: {outputFile}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo: {e}")

def abrir_pasta():
    diretorio_atual = os.getcwd()
    pasta_docs = os.path.join(diretorio_atual, "docs")
    
    # Verifica se a pasta 'docs' existe dentro do diretório atual
    if os.path.exists(pasta_docs) and os.path.isdir(pasta_docs):
        diretorio_para_abrir = pasta_docs
    else:
        diretorio_para_abrir = diretorio_atual
    
    subprocess.Popen(f'explorer {os.path.realpath(diretorio_para_abrir)}')

# Criar a janela principal
root = tk.Tk()
root.title("Conversor de Documentos | By: Sidnei Santiago")
root.geometry("526x203+300+150")
root.resizable(height=False, width=False)

# Criar um botão para selecionar o arquivo e converter
botao_converter = tk.Button(root, text="Selecionar PDF e Converter", command=converter_pdf_para_word)
botao_converter.pack(pady=20)

botao_converter2 = tk.Button(root, text="Selecionar WORD e Converter", command=converter_word_para_pdf)
botao_converter2.pack(pady=20)

botao_abrir_pasta = tk.Button(root, text="Abrir Pasta", command=abrir_pasta)
botao_abrir_pasta.pack(pady=20)

# Criar um label para exibir o caminho do arquivo selecionado
label = tk.Label(root, text="Nenhum arquivo selecionado")
label.pack(pady=20)

# Iniciar o loop da interface gráfica
root.mainloop()
