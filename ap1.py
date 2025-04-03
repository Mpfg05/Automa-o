import os
import pandas as pd
import pyautogui
import time
import webbrowser
from openpyxl import Workbook
from datetime import datetime 


tarefas_arquivo = "tarefas.csv"

if not os.path.exists(tarefas_arquivo):
    print(f"Arquivo NÃO encontrado: {tarefas_arquivo}")
    exit()

df = pd.read_csv(tarefas_arquivo)

df = df.dropna(subset=["Tarefa", "Tipo"])

tarefas = df.to_dict(orient="records")
relatorio = []

def executar_tarefa(tarefa):
    """Executa uma ação com base no tipo de tarefa e registra o resultado."""
    tipo = str(tarefa["Tipo"]).strip().lower()  # Normaliza para evitar erros
    dado = str(tarefa["Dado"]).strip()

    inicio = time.time() 
    
    try:
        if tipo == "abrir_navegador":
            webbrowser.open(dado)
        
        elif tipo == "click":
            partes = dado.split(",")
            if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
                x, y = map(int, partes)
                pyautogui.click(x, y)
            else:
                raise ValueError(f"Coordenadas inválidas: {dado}")
        
        elif tipo == "texto":
            pyautogui.write(dado)
        
        elif tipo == "tecla":
            pyautogui.press(dado)
        
        elif tipo == "espera":
            try:
                tempo = float(dado) if dado else 1 
                time.sleep(tempo)
            except ValueError:
                raise ValueError(f"Valor inválido para espera: {dado}")
        
        elif tipo == "scroll":
            try:
                pyautogui.scroll(int(dado))
            except ValueError:
                raise ValueError(f"Valor inválido para scroll: {dado}")

        else:
            raise ValueError(f"Tipo desconhecido: {tipo}")
        
        status = "Sucesso"
    
    except Exception as e:
        status = f"Erro: {str(e)}"
    
    fim = time.time()  
    tempo_execucao = round(fim - inicio, 2)  
    
    relatorio.append({"Tarefa": tarefa["Tarefa"], "Status": status, "Tempo (s)": tempo_execucao})

print("Abrindo o navegador...")
time.sleep(2)
pyautogui.press("win")
pyautogui.write("Chrome", interval=0.2)
time.sleep(1)
pyautogui.press("enter")
time.sleep(2)
webbrowser.open("https://www.google.com")
time.sleep(2)
pyautogui.write("www.youtube.com", interval=0.1)
time.sleep(2)
pyautogui.press("enter")
time.sleep(2)

x_pagina, y_pagina = 379, 297  
pyautogui.click(x_pagina, y_pagina)
time.sleep(4)

x_pesquisa, y_pesquisa = 1234, 121  
pyautogui.click(x_pesquisa, y_pesquisa)
time.sleep(2)

pyautogui.write("Faculdade Impacta", interval=0.1)
time.sleep(1)
pyautogui.press("enter")
time.sleep(2)
pyautogui.scroll(-400)
time.sleep(2)

x_video, y_video = 1113, 773 
time.sleep(2)
pyautogui.click(x_video, y_video)
time.sleep(2)

pyautogui.alert("Automação finalizada! Começando processo de relatório...")
time.sleep(1)


for tarefa in tarefas:
    print(f"Executando: {tarefa['Tarefa']}")
    executar_tarefa(tarefa)


wb = Workbook()
ws = wb.active
ws.title = "Relatório de Execução"
ws.append(["Tarefa", "Status", "Tempo (s)"])

for linha in relatorio:
    ws.append([linha["Tarefa"], linha["Status"], linha["Tempo (s)"]])

relatorio_arquivo = datetime.now().strftime("relatorio_execucao%Y%m%d_%H%M%S.xlsx")
wb.save(relatorio_arquivo)
print(f"Relatório salvo em: {relatorio_arquivo}")
