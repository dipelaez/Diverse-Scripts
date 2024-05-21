"""
Este script lida com finais de semana, acessa um banco de dados, extrai um relatório da Ploomes,
processa o relatório e insere os dados no banco de dados.

Dependências:
    - datetime
    - time
    - pyodbc
    - sqlalchemy
    - selenium
    - pyautogui
    - pandas
    - os

Como usar:
    python script_interacoes_ploomes.py
    É necessário estar com o windows livre para o pyautogui

Este script faz:
    - Ajusta a data para a última sexta-feira se hoje for segunda-feira
    - Conecta ao banco de dados
    - Baixa um relatório da Ploomes
    - Processa o relatório e insere dados no banco de dados

Criado por: Diego Nunes
Data: 21-05-2024
"""

import os
from time import sleep
from datetime import datetime, timedelta
import urllib
import pandas as pd
from sqlalchemy import create_engine
from selenium import webdriver
import pyautogui

def ajustar_data_final_de_semana():
    """
    Ajusta a data para a última sexta-feira se hoje for segunda-feira,
    caso contrário, usa a data atual.

    Retorna:
        str: A data ajustada como string no formato 'YYYY-MM-DD'.
    """
    now = datetime.now()
    day_of_week = now.strftime("%A")

    if day_of_week == 'Monday':
        friday = now - timedelta(days=3)
        current_date_str = friday.strftime("%Y-%m-%d")
    else:
        current_date_str = now.strftime("%Y-%m-%d")

    return current_date_str

def conectar_banco():
    """
    Conecta ao banco de dados usando as credenciais fornecidas e retorna a engine de conexão.

    Retorna:
        engine: A engine de conexão do SQLAlchemy.
    """
    print('Acessando Banco de Dados')
    username = 'ploomesInteracoes'
    password = 'XXXXXXX' # Editar aqui
    server = 'xxxxx' # Editar
    database = 'xxxxx' # Editar

    connection_string = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
    connection_url = f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
    engine = create_engine(connection_url)
    print('Acesso ao Banco OK!')
    
    return engine

def baixar_relatorio_ploomes():
    """
    Baixa o relatório da Ploomes usando automação de navegador e salva o arquivo na pasta Downloads.
    """
    print('Baixando relatório da Ploomes')
    driver = webdriver.Chrome()

    service = webdriver.ChromeService(log_output='log_path.txt')

    driver.get('https://app10.ploomes.com/')
    sleep(5)

    # Inserindo email de login
    pyautogui.write('user@user.com.br')
    sleep(1)
    pyautogui.press('tab')
    sleep(1)

    # Inserindo senha de login
    pyautogui.write('password')
    sleep(1)
    pyautogui.press('enter')
    sleep(20)

    pyautogui.leftClick(750,171)
    sleep(3)
    pyautogui.leftClick(421,373)
    sleep(3)
    pyautogui.leftClick(340,508)
    sleep(2)

    # Clica em exportar relatório
    pyautogui.leftClick(791,664)
    sleep(5)

    # Clica em popup "Atenção! 10 planilhas"
    pyautogui.leftClick(587,383)
    sleep(10)

    # Aguarda imagem aparecer
    pyautogui.locateOnScreen("exportar2.png", minSearchTime=1000)
    pyautogui.click('exportar2.png')
    sleep(5)
    pyautogui.leftClick(820,624)
    sleep(5)

    driver.close()
    print('Relatório baixado!')

def processar_e_inserir_dados(engine, current_date_str):
    """
    Processa o relatório baixado e insere os dados no banco de dados.

    Args:
        engine: A engine de conexão do SQLAlchemy.
        current_date_str: A data atual ou ajustada como string no formato 'YYYY-MM-DD'.
    """
    print('Limpando relatório e inserindo dados no banco!')
    df = pd.read_excel("User/Downloads/RelatorioDiarioInteracoes.xlsx",
                       converters={'Código ERP do Criador': str, 'Código ERP do Cliente': str, 'Código da Proposta do Negócio': str})
    df['DtCriado'] = (datetime.now()).strftime('%Y%m%d')
    df = df.fillna('')
    df.rename(columns={'Código ERP do Criador': 'CodERPCriador', 'Código ERP do Cliente': 'CodERPCliente', 'Descrição': 'Descricao',
                       'Código da Proposta do Negócio': 'Proposta'}, inplace=True)
    os.remove("User/Downloads/RelatorioDiarioInteracoes.xlsx")

    # Selecionando apenas um dia
    df_limpo = df.query("Data == @current_date_str")

    # Adicionando dados à tabela do banco
    df_limpo.to_sql('BD_DADOS_PLOOMES_INTERACOES', con=engine, if_exists='append', index=False)
    print('Dados inseridos!')

def main():
    """
    Função principal que executa o fluxo completo do script.
    """
    current_date_str = ajustar_data_final_de_semana()
    engine = conectar_banco()
    baixar_relatorio_ploomes()
    processar_e_inserir_dados(engine, current_date_str)

if __name__ == "__main__":
    main()
