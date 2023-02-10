from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from numpy import random
from mensagem import *
import time
import datetime
import urllib
import pandas as pd
import os
import PySimpleGUI as sg
import ctypes

ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)

sg.theme('lightGrey6')

element = [
    [sg.Text("Para qual tipo de cliente gostaria de enviar?",font=('Roboto', 12))],
]

element1 = [
    [sg.HorizontalSeparator()],
    [sg.VerticalSeparator(),sg.Radio('eCPF', 'group 1', key='cpf', enable_events=True), sg.Radio('eCNPJ', 'group 1', key='cnpj', enable_events=True),sg.VerticalSeparator()],
    [sg.HorizontalSeparator()],
]

element3 = [
    [sg.Button('Enviar', font=('Roboto', 10))],
]

layout = [
    [sg.Column(element)],
    [sg.Column(element1)],
    [sg.Column(element3)],
]

janela = sg.Window('Sistema de ponto WebCertificados', layout,element_justification='c')

def isNaN(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return False

with open("Debug.txt", "a", encoding="utf-8") as arquivo:
    data = datetime.datetime.now()
    date_time = data.strftime("%d/%m/%Y, %H:%M")
    arquivo.write(f"\n---------------------------------------------------------------------------------------------------------\nComeço do disparo {date_time}\n\n")

def sendEcpf(localPlan,limiteQuantidade):
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://web.whatsapp.com/")

    while len(navegador.find_elements(By.ID, value='side')) < 1:
        time.sleep(1)

    contatos = pd.read_excel(localPlan)
    quantidade = 0

    
    for i, mensagem in enumerate(contatos["Tel Celular"]):
        modelo = contatos.loc[i, "Produto"]
        nome = contatos.loc[i, "Representante Nome"]
        data = contatos.loc[i, "Data de vencimento"]
        protocolo = contatos.loc[i, "PROTOCOLO"]
        numero = contatos.loc[i, "Tel Celular"]
        
        texto = urllib.parse.quote(f"{fisica}"%(modelo, nome, protocolo, data))
        if numero != "nan":
            if quantidade < int(limiteQuantidade):
                    link = f"https://web.whatsapp.com/send?phone=55{numero}&text={texto}"
                    navegador.get(link)

                    while len(navegador.find_elements(By.ID, value='side')) < 1:
                        time.sleep(1)

                    time.sleep(5)
                    with open("Debug.txt", "a", encoding="utf-8") as arquivo:
                        if len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
                            while not(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')):
                                time.sleep(1)
                            
                            if navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span'):
                                navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
                                quantidade += 1
                                arquivo.write(f"{quantidade} {numero} Sucesso\n")
                                numeroRand = random.randint(10,15)
                                time.sleep(numeroRand)
                            else:
                                arquivo.write(f"{numero} Problematico\n")
                        else:
                            arquivo.write(f"{numero} Problematico\n")
            else:
                return sg.popup(f"Maximo de envios atingido!\nEnviei no total {quantidade} de mensagens com sucesso!")

    with open("Debug.txt", "a", encoding="utf-8") as arquivo:
        data = datetime.datetime.now()
        date = data.strftime("%d/%m/%Y")
        hour = data.strftime("%H:%M")
        arquivo.write(f"\n\nTermino do disparo do dia {date} às {hour}\n---------------------------------------------------------------------------------------------------------\n")
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000) #set the setting back to normal

def sendEcnpj(localPlan,limiteQuantidade):
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://web.whatsapp.com/")

    while len(navegador.find_elements(By.ID, value='side')) < 1:
        time.sleep(1)

    contatos = pd.read_excel(localPlan)
    quantidade = 0

    
    for i, mensagem in enumerate(contatos["Tel Celular"]):
        modelo = contatos.loc[i, "Produto"]
        nome = contatos.loc[i, "Representante Nome"]
        data = contatos.loc[i, "Data de vencimento"]
        protocolo = contatos.loc[i, "PROTOCOLO"]
        nempresa = contatos.loc[i, "Cliente Nome"]
        numero = contatos.loc[i, "Tel Celular"]
        
        texto = urllib.parse.quote(f"{empresa}"%(modelo, nempresa, protocolo, data))
        if numero != "nan":
            if quantidade < int(limiteQuantidade):
                    link = f"https://web.whatsapp.com/send?phone=55{numero}&text={texto}"
                    navegador.get(link)

                    while len(navegador.find_elements(By.ID, value='side')) < 1:
                        time.sleep(1)

                    time.sleep(5)
                    with open("Debug.txt", "a", encoding="utf-8") as arquivo:
                        if len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
                            while not(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')):
                                time.sleep(1)
                            
                            if navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span'):
                                navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
                                quantidade += 1
                                arquivo.write(f"{quantidade} {numero} Sucesso\n")
                                numeroRand = random.randint(10,15)
                                time.sleep(numeroRand)
                            else:
                                arquivo.write(f"{numero} Problematico\n")
                        else:
                            arquivo.write(f"{numero} Problematico\n")
            else:
                return sg.popup(f"Maximo de envios atingido!\nEnviei no total {quantidade} de mensagens com sucesso!")

    with open("Debug.txt", "a", encoding="utf-8") as arquivo:
        data = datetime.datetime.now()
        date = data.strftime("%d/%m/%Y")
        hour = data.strftime("%H:%M")
        arquivo.write(f"\n\nTermino do disparo do dia {date} às {hour}\n---------------------------------------------------------------------------------------------------------\n")
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000) #set the setting back to normal

while True:
    eventos,valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000) #set the setting back to normal
        break
    if eventos == 'Enviar':
        arquivo = sg.popup_get_file('Selecione a planilha:')
        if arquivo:
            numeros = sg.popup_get_text('para quantos numeros gostaria de enviar?')
            if numeros:
                if valores["cpf"]:
                    sendEcpf(arquivo,numeros)
                elif valores["cnpj"]:
                    sendEcnpj(arquivo,numeros)
                else:
                    sg.popup("Nenhuma opção escolhida")