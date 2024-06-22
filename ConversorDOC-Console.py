from time import sleep
from pdf2docx import Converter
import os
import win32com.client

def converter_word_para_pdf():
    # Determinar o diretório onde o script Python está localizado
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo de entrada
    inputFile = os.path.join(diretorio_atual, "teste.docx")

    # Nome do arquivo sem extensão
    file_name = os.path.splitext(os.path.basename(inputFile))[0]

    # Verificar se a extensão do arquivo é .docx
    extensao = os.path.splitext(inputFile)[1]
    if extensao.lower() != ".docx":
        print("Erro: Por favor, selecione um arquivo Word (.docx).")
        sleep(2)
        menu_principal()

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
        doc.SaveAs(outputFile, FileFormat=17)  # 17 corresponde ao formato PDF
        doc.Close()
        word.Quit()
        print(f"Sucesso! Arquivo convertido com sucesso e salvo em: {outputFile}")
        sleep(4)
        menu_principal()
    except Exception as e:
        print(f"Erro ao converter o arquivo: {e}")
        sleep(4)
        menu_principal()


def converter_pdf_para_word():
    # Determinar o diretório onde o script Python está localizado
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo de entrada
    inputFile = os.path.join(diretorio_atual, "image.png")

    # Verificar se a extensão do arquivo é .pdf
    extensao = os.path.splitext(inputFile)[1]
    if extensao.lower() != ".pdf":
        print("ERROR!! Por favor, selecione um arquivo PDF.")
        sleep(2)
        menu_principal()
    
    # Criar uma pasta específica dentro do diretório atual
    pasta_destino = os.path.join(diretorio_atual, "docs")
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    # Nome do arquivo de saída
    nome_arquivo_saida = os.path.splitext(os.path.basename(inputFile))[0] + ".docx"
    
    # Caminho completo para o arquivo de saída dentro da pasta especificada
    caminho_saida = os.path.join(pasta_destino, nome_arquivo_saida)
    
    # Converter PDF para DOCX
    try:
        cv = Converter(inputFile)
        cv.convert(caminho_saida, start=0, end=None) 
        cv.close()
        print(f"Sucesso!! Arquivo convertido e salvo em: {caminho_saida}")
        sleep(4)
        menu_principal()
    except Exception as e:
        print(f"ERROR! Erro ao converter o arquivo: {e}")
        sleep(4)
        menu_principal()


#-------------------------------------- MENU PRINCIPAL ---------------------------------

def menu_principal():
    print("Escolha qual tipo de conversão deseja fazer: ")

    print("(1) Word para PDF \n(2) PDF para Word \n(0) Sair");

    opcao = int(input())

    if opcao == 1:
        converter_word_para_pdf()
    elif opcao == 2:
        converter_pdf_para_word()
    else:
        return
    
menu_principal()
