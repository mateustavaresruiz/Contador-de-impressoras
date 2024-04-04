from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from pygetwindow import getWindowsWithTitle
import os
import sqlite3
import tkinter as tk
from tkinter import ttk

# BANCO DE DADOS SQLITE3

def conectar_bd():
    # Conecta ao banco de dados SQLite e cria a tabela 'impressoras' se não existir
    conexao = sqlite3.connect('impressoras.db')
    cursor = conexao.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS impressoras(
            id INTEGER PRIMARY KEY,
            nome TEXT,
            endereco_ip TEXT,
            selector TEXT
        )
    ''')
    conexao.commit()
    return conexao, cursor

def adicionar_impressora():
    # Obtém informações da interface do usuário e insere na tabela
    nome = nome_entry.get()
    endereco_ip = endereco_ip_entry.get()
    selector = selector_entry.get()
    cursor.execute('INSERT INTO impressoras (nome, endereco_ip, selector) VALUES (?, ?, ?)',
                   (nome, endereco_ip, selector))
    conexao.commit()
    atualizar_tabela()

def editar_impressora():
    # Obtém os valores da impressora selecionada e preenche os campos de edição
    item_selecionado = tabela.selection()[0]
    valores = tabela.item(item_selecionado, 'values')
    nome_entry.delete(0, tk.END)
    nome_entry.insert(0, valores[1])
    endereco_ip_entry.delete(0, tk.END)
    endereco_ip_entry.insert(0, valores[2])
    selector_entry.delete(0, tk.END)
    selector_entry.insert(0, valores[3])

def excluir_impressora():
    # Exclui a impressora selecionada da tabela
    selected_item = tabela.focus()
    values = tabela.item(selected_item, 'values')
    id = values[0]
    cursor.execute('DELETE FROM impressoras WHERE id=?', (id,))
    conexao.commit()
    atualizar_tabela()

def atualizar_tabela():
    # Atualiza a tabela exibindo todas as impressoras
    tabela.delete(*tabela.get_children())
    cursor.execute('SELECT * FROM impressoras')
    for row in cursor.fetchall():
        tabela.insert('', 'end', values=row)

def limpar():
    # Limpa os campos de entrada
    nome_entry.delete(0, tk.END)
    endereco_ip_entry.delete(0, tk.END)
    selector_entry.delete(0, tk.END)

def salvar():
    # Salva as alterações feitas na impressora selecionada
    selected_item = tabela.focus()
    values = tabela.item(selected_item, 'values')
    id = values[0]
    nome = nome_entry.get()
    endereco_ip = endereco_ip_entry.get()
    selector = selector_entry.get()
    cursor.execute('UPDATE impressoras SET nome=?, endereco_ip=?, selector=? WHERE id=?',
                   (nome, endereco_ip, selector, id))
    conexao.commit()
    atualizar_tabela()

# CODIGO QUE EXECUTA O NAVEGADOR 

def executar():
    # Cria uma pasta na área de trabalho para armazenar as capturas de página
    pasta_area_de_trabalho = os.path.join(os.path.expanduser('~'), 'Desktop', 'Capturas_Pagina')
    os.makedirs(pasta_area_de_trabalho, exist_ok=True)

    # Configurações do navegador Chrome
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--log-level=3")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-logging")
    options.add_argument("--window-size=1300,1000")
    options.add_argument("--headless")
    prefs = {
        "profile.managed_default_content_settings.stylesheet": 2,
        "profile.default_content_setting_values.notifications": 2,
    }
    options.add_experimental_option("prefs", prefs)

    # Inicializa o driver do Chrome
    driver = webdriver.Chrome(options=options)

    def obter_todas_impressoras():
        # Conecta ao banco de dados SQLite para obter informações sobre as impressoras
        conexao = sqlite3.connect('impressoras.db')
        cursor = conexao.cursor()
        cursor.execute('SELECT * FROM impressoras')
        dados = cursor.fetchall()
        conexao.close()
        return dados

    def obter_conteudo_impressora(nome_impressora, endereco_pagina, seletor_css, arquivo_saida, driver):
        try:
            print(f"Processando impressora: {nome_impressora}")

            # Acessa a página da impressora
            driver.get(endereco_pagina)

            # Aguarda até que o elemento especificado pelo seletor CSS esteja visível
            elemento = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, seletor_css))
            )

            # Obtém o conteúdo do elemento e escreve no arquivo de saída
            conteudo = elemento.text
            with open(arquivo_saida, "a", encoding="utf-8") as arquivo:
                arquivo.write(f"{nome_impressora} : {conteudo}\n")

            # Captura um screenshot da página
            captura_screenshot = os.path.join(pasta_area_de_trabalho, f"{nome_impressora}.png")
            driver.save_screenshot(captura_screenshot)

        except (TimeoutException, NoSuchElementException, WebDriverException) as e:
            mensagem_erro = f"{nome_impressora} : Erro! Verificar se a impressora está ligada\n"
            print(mensagem_erro)

            with open(arquivo_saida, "a", encoding="utf-8") as arquivo:
                arquivo.write(mensagem_erro)

    def capturar_pagina_inteira(driver, file_path):
        # Captura a página inteira (incluindo a parte não visível)
        altura_total = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")

        driver.set_window_size(driver.get_window_size().get("width"), altura_total)
        driver.set_window_position(0, driver.execute_script("return window.screen.availHeight - window.outerHeight;"))
        driver.save_screenshot(file_path)

    if __name__ == "__main__":
        impressoras = obter_todas_impressoras()

        # Arquivo de saída para registrar os resultados
        arquivo_saida = os.path.join(os.path.expanduser('~'), 'Desktop', 'contador_impressoras.txt')

        # Processa cada impressora
        for impressora in impressoras:
            nome_impressora, endereco_pagina, seletor_css = impressora[1], impressora[2], impressora[3]
            obter_conteudo_impressora(nome_impressora, endereco_pagina, seletor_css, arquivo_saida, driver)

        # Encerra o driver do Chrome
        driver.quit()

# Cria a janela principal
root = tk.Tk()
root.title('Gerenciador de Impressoras')

input_frame = ttk.Frame(root, padding="20")
input_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

ttk.Label(input_frame, text='Nome da Impressora:').grid(row=0, column=0, sticky="w")
nome_entry = ttk.Entry(input_frame, width=45)
nome_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

ttk.Label(input_frame, text='Endereço IP:').grid(row=1, column=0, sticky="w")
endereco_ip_entry = ttk.Entry(input_frame, width=45)
endereco_ip_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

ttk.Label(input_frame, text='Selector:').grid(row=2, column=0, sticky="w")
selector_entry = ttk.Entry(input_frame, width=45)
selector_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

button_frame = ttk.Frame(root, padding="10")
button_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

btn_adicionar = ttk.Button(button_frame, text='Adicionar', command=adicionar_impressora)
btn_adicionar.grid(row=0, column=0, padx=5, pady=5)

btn_editar = ttk.Button(button_frame, text='Editar', command=editar_impressora)
btn_editar.grid(row=0, column=1, padx=5, pady=5)

btn_excluir = ttk.Button(button_frame, text='Excluir', command=excluir_impressora)
btn_excluir.grid(row=0, column=2, padx=5, pady=5)

btn_limpar = ttk.Button(button_frame, text='Limpar', command=limpar)
btn_limpar.grid(row=0, column=3, padx=5, pady=5)

btn_executar = ttk.Button(button_frame, text='Coletar', command=executar)
btn_executar.grid(row=0, column=4, padx=5, pady=5)

btn_salvar = ttk.Button(input_frame, text='Salvar', command=salvar)
btn_salvar.grid(row=1, column=4, padx=30, pady=5, sticky="e")

table_frame = ttk.Frame(root, padding="20")
table_frame.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")

table_frame = ttk.Frame(root, padding="20")
table_frame.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")

table_frame = ttk.Frame(root, padding="20")
table_frame.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")

scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal")
scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical")

# Crie a tabela
tabela = ttk.Treeview(table_frame, columns=('ID', 'Nome da Impressora', 'Endereço IP', 'Selector'), show='headings', xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)
tabela.heading('ID', text='ID')
tabela.heading('Nome da Impressora', text='Nome da Impressora')
tabela.heading('Endereço IP', text='Endereço IP')
tabela.heading('Selector', text='Selector')

# Configure as barras de rolagem
scrollbar_x.config(command=tabela.xview)
scrollbar_y.config(command=tabela.yview)

tabela.grid(row=0, column=0, sticky="nsew")
scrollbar_x.grid(row=1, column=0, sticky="ew")
scrollbar_y.grid(row=0, column=1, sticky="ns")

table_frame.rowconfigure(0, weight=1)
table_frame.columnconfigure(0, weight=1)

table_frame.rowconfigure(0, weight=1)
table_frame.columnconfigure(0, weight=1)

conexao, cursor = conectar_bd()
atualizar_tabela()

root.mainloop()
