import tkinter as tk
from tkinter import filedialog
import webbrowser
from urllib.parse import quote
import time
import os

import pyautogui
import openpyxl

def selecionar_arquivo(tipo_de_arquivo: str) -> str:
    """
    Função que pede para selecionar o .txt e o .xlsx
    param: tipo_de_arquivo - txt | xlsx
    """
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    if tipo_de_arquivo == '.txt':
        filtro = [('Text files', '*.txt')]
    elif tipo_de_arquivo == '.xlsx':
        filtro = [('Excel files', '*.xlsx')]
    else:
        raise ValueError("Tipo de arquivo inválido. Use 'txt' ou 'xlsx'.")

    arquivo = filedialog.askopenfilename(title=f"Selecione um arquivo {tipo_de_arquivo}", filetypes=filtro)

    if arquivo:
        print(f"Arquivo selecionado: {arquivo}")
    else:
        print("Nenhum arquivo selecionado.")
    return arquivo

def get_mensagem_de_texto(texto_path: str) -> str:
    """
    Lê o conteúdo de um arquivo de texto.
    """
    with open(texto_path, 'r', encoding='utf-8') as arquivo_de_texto:
        mensagem_de_texto = arquivo_de_texto.read()
    return mensagem_de_texto

# Main TODO: Ceritificar que user tá logado no wpp web
if __name__ == '__main__':
    # Seleciona o arquivo de texto para a mensagem
    texto_path: str = selecionar_arquivo('.txt')
    mensagem: str = get_mensagem_de_texto(texto_path=texto_path)

    print(f'{mensagem=}\n')
    
    planilha_path: str = selecionar_arquivo('.xlsx')
    
    workbook = openpyxl.load_workbook(planilha_path)
    planilha_clientes = workbook[workbook.sheetnames[0]] # 
    # TODO:  deixar a seleção por str ['nome_col']
    for linha in planilha_clientes.iter_rows(min_row=2):
        NOME = linha[0].value
        TELEFONE = linha[1].value

        if not NOME or not TELEFONE:  # Ignora linhas incompletas
            continue

        try:
            # Abre o WhatsApp Web com o link formatado
            webbrowser.open(f'https://web.whatsapp.com/send?phone={TELEFONE}&text={quote(mensagem)}')
            time.sleep(10)  # Aguarda o envio da mensagem
            pyautogui.hotkey('enter')  # Fecha a aba do navegador
            time.sleep(5)  # Aguarda o envio da mensagem
            pyautogui.hotkey('ctrl', 'w')  # Fecha a aba do navegador
        except Exception as e:
            print(f'Erro ao enviar mensagem para {NOME}: {e}')
            # Salva o erro em um arquivo
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{NOME},{TELEFONE}{os.linesep}')

    print("Envio de mensagens concluído!")
