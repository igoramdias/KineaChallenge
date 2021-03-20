from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os
import pandas as pd
import numpy as np
from datetime import date 

def date_to_file(dt: str, type: str) -> str:
    """
        Função para converter a data em nome dos files

        :param dt: Data a ser analisada
        :param type: Para qual documento será feito o ajuste
    """

    dt = dt.split('/') # Retira os chars /

    # Criação dos nome dos files de acordo com a data o tipo passado
    if type == 'TAXAS': 
        mes = {1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun', 7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'}
        file = 'd'+dt[2][-2:]+mes[int(dt[1])]+dt[0]+'.xls'

    if type == 'REUNE':
        file = 'REUNE_Acumulada_'+''.join(dt)+'.csv' 

    if type == 'IMAB':
        file = 'IMA_'+''.join(dt)+'.csv'

    if type == 'ETTJ':
        file = 'CurvaZero_'+''.join(dt)+'.csv'
    
    return file # Retorna o nome do file

def switch_day_month(dt: str):
    """
        Troca as posições do mês com o dia na data
    """

    new_dt = dt[3:5] + '/' + dt[0:2] + '/' + dt[6:10]

    return new_dt

def crawl_data_week(star_date: str, end_date: str) -> None:
    """
        Função para realizar a coleta de dados de REUNE, IMAB, ETTJ e TAXAS

        :param star_date: Dia inicial (útil ou não) ao qual começará a pedir os dados
        :param end_date: Dia final (útil ou não) ao qual terminará a pedir os dados
    """

    sources = {
        'IMAB': "https://www.anbima.com.br/pt_br/informar/ima-resultados-diarios.htm",
        'REUNE': "https://www.anbima.com.br/pt_br/informar/sistema-reune.htm",
        'ETTJ': "https://www.anbima.com.br/pt_br/informar/curvas-de-juros-fechamento.htm", 
        'TAXAS': "https://www.anbima.com.br/informacoes/merc-sec-debentures/default.asp"
    }

    downloads_path = "~\Downloads"
    downloads_path = os.path.expanduser(downloads_path)

    chrome_options = Options()

    driver = webdriver.Chrome(
        executable_path=os.path.join(downloads_path, "chromedriver.exe"), options=chrome_options
    )

    Feriados = list(pd.read_excel(os.path.join(downloads_path, "Feriados.xlsx"), header=None)[0])
    Feriados = [str(d.date()) for d in Feriados]

    star_date = switch_day_month(star_date)
    end_date = switch_day_month(end_date)
    list_dt = [date.strftime(dt, '%d/%m/%Y') for dt in pd.bdate_range(start=star_date, end=end_date, freq='C', holidays=Feriados)]

    # Fazer um loop de datas aqui
    for dia in list_dt:
        
        new_dir = os.path.join(downloads_path, "{}".format(dia.replace('/', '')))
        if not (dia.replace('/','') in [file for file in os.listdir(downloads_path)]): 
            os.mkdir(new_dir)

        for src in sources:

            if np.busday_count(pd.to_datetime(dia, dayfirst=True).date(), pd.to_datetime(date.today()).date()) <= 5:
                if src in ['IMAB', 'TAXAS', 'ETTJ']:
                    driver.get(sources[src])
                    if src == 'IMAB':
                        print('Acessando site do IMAB para o dia {}...'.format(dia))

                        WebDriverWait(driver, 10).until(
                            EC.frame_to_be_available_and_switch_to_it((By.CLASS_NAME,"full"))
                        )
                        
                        driver.execute_script("document.getElementsByName('escolha')[1].click()") # Define tipo de visualização
                        driver.execute_script("document.getElementsByName('saida')[1].click()") # Define tipo de arquivo 
                        driver.execute_script("document.getElementsByName('indice')[4].click()") # Define indice de consulta
                        driver.execute_script("document.getElementsByName('consulta')[1].click()") # Define tipo de consulta
                        driver.execute_script("document.getElementsByName('Dt_Ref')[0].value = '{}'".format(dia)) # Define tipo de arquivo
                        driver.execute_script("document.getElementsByName('Consultar')[0].click()") # Realiza a consulta
                        
                        ant_file = os.path.join(downloads_path, date_to_file(dia, 'IMAB'))
                        new_file = os.path.join(new_dir, date_to_file(dia, 'IMAB'))

                        print('Coleta concluída!')
                    
                    elif src == 'TAXAS':
                        print('Acessando site de TAXAS para o dia {}...'.format(dia))

                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.NAME,"Dt_Ref"))
                        )
                        
                        driver.execute_script("document.getElementsByName('Dt_Ref')[0].value = '{}'".format(dia)) # Define tipo de arquivo
                        driver.execute_script("document.getElementsByName('Consultar')[0].click()") # Realiza a consulta
                        
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME,"linkinterno"))
                        )

                        driver.execute_script("document.getElementsByClassName('linkinterno')[2].click()") # Realiza a consulta
                        
                        ant_file = os.path.join(downloads_path, date_to_file(dia, 'TAXAS'))
                        new_file = os.path.join(new_dir, date_to_file(dia, 'TAXAS'))

                        print('Coleta concluída!')
                    
                    elif src == 'ETTJ':
                        print('Acessando site do ETTJ para o dia {}...'.format(dia))

                        WebDriverWait(driver, 10).until(
                            EC.frame_to_be_available_and_switch_to_it((By.CLASS_NAME,"full"))
                        )
                        
                        driver.execute_script("document.getElementsByName('escolha')[1].click()") # Define tipo de visualização
                        driver.execute_script("document.getElementsByName('saida')[1].click()") # Define tipo de arquivo 
                        driver.execute_script("document.getElementsByName('Dt_Ref')[0].value = '{}'".format(dia)) # Define tipo de arquivo
                        driver.execute_script("document.getElementsByName('Consultar')[0].click()") # Realiza a consulta
                        
                        ant_file = os.path.join(downloads_path, date_to_file(dia, 'ETTJ'))
                        new_file = os.path.join(new_dir, date_to_file(dia, 'ETTJ'))

                        print('Coleta concluída!')
                
                    while not os.path.exists(ant_file):
                        sleep(1)

                    if os.path.isfile(ant_file):
                        os.replace(ant_file, new_file)
                    else:
                        raise ValueError("%s isn't a file!" % ant_file)
                else:
                    pass

            if src == 'REUNE':
                driver.get(sources[src])

                print('Acessando site do REUNE para dia {}...'.format(dia))
                
                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.CLASS_NAME,"full"))
                )
                
                driver.execute_script("document.getElementsByName('Dt_Ref')[0].value = '{}'".format(dia)) # Define a data
                driver.execute_script("document.getElementById('TpInstFinanceiro').value = 'DEB'") # Define tipo de instrumento
                driver.execute_script("document.getElementsByName('escolha')[1].click()") # Define tipo de visualização
                driver.execute_script("document.getElementsByName('saida')[1].click()") # Define tipo de arquivo 
                driver.execute_script("document.getElementsByName('Consultar')[0].click()") # Realiza a consulta
                    
                ant_file = os.path.join(downloads_path, date_to_file(dia, 'REUNE'))
                new_file = os.path.join(new_dir, date_to_file(dia, 'REUNE'))

                print('Coleta concluída!')
        
                while not os.path.exists(ant_file):
                    sleep(1)

                if os.path.isfile(ant_file):
                    os.replace(ant_file, new_file)
                else:
                    raise ValueError("%s isn't a file!" % ant_file)

    driver.close()

def crawl_data_today() -> None:
    """
        Função para realizar a coleta de dados de DEB
    """
    
    source = "http://www.debentures.com.br/exploreosnd/consultaadados/emissoesdedebentures/caracteristicas_f.asp?tip_deb=publicas"

    downloads_path = "~\Downloads"
    downloads_path = os.path.expanduser(downloads_path)
    
    chrome_options = Options()

    driver = webdriver.Chrome(
        executable_path=os.path.join(downloads_path, "chromedriver.exe"), options=chrome_options
    )
    
    dia = date.strftime(date.today(), '%d/%m/%Y')
    
    new_dir = os.path.join(downloads_path, "{}".format(dia.replace('/', '')))
    if not (dia.replace('/','') in [file for file in os.listdir(downloads_path)]): 
        os.mkdir(new_dir)
    
    driver.get(source)

    print('Acessando site de Debênture.com.br para dia {}...'.format(dia))
    
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME,"Ver11FF6600"))
    )
    
    driver.execute_script("document.getElementsByClassName('Ver11FF6600')[0].click()") # Realiza a consulta

    while len([file for file in os.listdir(downloads_path) if 'Debentures.com.br' in file]) == 0:
        sleep(10)
   
    print('Coleta concluída!')
    
    DATA_DEB_path = str([file for file in os.listdir(downloads_path) if 'Debentures' in file][0])
    ant_file = os.path.join(downloads_path, DATA_DEB_path)
    new_file = os.path.join(new_dir, DATA_DEB_path)

    if os.path.isfile(ant_file):
        os.replace(ant_file, new_file)
    else:
        raise ValueError("%s isn't a file!" % ant_file)

    driver.close()