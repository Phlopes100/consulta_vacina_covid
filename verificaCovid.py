from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import pyautogui
from openpyxl import Workbook


class ConsultaVacina:

    def __init__(self):
        self.site_consulta = 'https://vacinacao.sms.fortaleza.ce.gov.br/pesquisa/atendidos'
        #aqui voce coloca o caminho do arquivo da listagem.xlsx
        self.listagem_covid = pd.read_excel(r"C:\Users\Pedro TI\Desktop\ProjetosFretcar\listagemCovid\listagem.xlsx")
        s=Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=s)

    def inicializador(self):
        self.navegacao()

    def navegacao(self):
        self.driver.get(self.site_consulta)
        
        print("Aguardando o carregamento da página...")
        time.sleep(15)

        #aqui eu fiz uma contagem de quantidade de cpf para percorrer
        for i in range(len(self.listagem_covid['CPF'])):
            print('Verificando... ')

            #o arquivo listagem tem 2 colunas NOME e CPF
            
            nome = self.listagem_covid['NOME'][i]
            cpf = self.listagem_covid['CPF'][i]
            verifica_cpf = len(f'{cpf}')

            #aqui removi o spin de consulta para não atrapalhar no processo.
            self.driver.execute_script(
                "$(document).ready(function(){$('#btnBuscar').click(function(){$('#cover-spin').remove();});});")
            pyautogui.press('enter')

            #aqui ajustei o cpf por conta dos zeros
            if verifica_cpf == 10:
                print("Vamos lá...")
                cpf = f'0{cpf}'
            elif verifica_cpf == 9:
                print("Vai dar certo...")
                cpf = f'00{cpf}'
            elif verifica_cpf == 8:
                print("Tudo vai acabar bem...")
                cpf = f'000{cpf}'

            campo_cpf = self.driver.find_element(By.XPATH, '//*[@id="cpf"]')
            campo_cpf.click()
            time.sleep(1)
            pyautogui.write(f'{cpf}', interval=0.05)
            self.driver.execute_script(
                "$(document).ready(function(){$('#btnBuscar').click(function(){$('#cover-spin').remove();});});")
            pyautogui.press('enter')
            time.sleep(2)

            try:
                self.driver.find_element(By.ID, "row_undefined")
                row_count = len(self.driver.find_elements(By.XPATH, '//*[@id="row_undefined"]/td[5]'))
                print('VACINAS SALVAM ' * 3)
                print(f'{nome} - TOMOU {row_count} DOSE(S)')
                print('-' * 20)
                campo_cpf.clear()
                time.sleep(1)

            except NoSuchElementException:
                print("-" * 20)
                print("-NENHUM REGISTRO ENCONTRADO-")
                print(f'{nome} {"TOMOU NENHUMA DOSE"}')
                print("-" * 20)
                campo_cpf.clear()
                time.sleep(1)
        
bot_covid = ConsultaVacina()
bot_covid.inicializador()


