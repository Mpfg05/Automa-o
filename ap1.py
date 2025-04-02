import os
import pandas as pd
import pyautogui
import time
import webbrowser
from openpyxl import Workbook


arquivo = r"C:\Users\Pichau\Desktop\faculdade\automação\tarefas.csv"  #verificar qual o caminho do csv


if not os.path.exists(arquivo):
    print(f"Arquivo NÃO encontrado: {arquivo}")
    exit()


df = pd.read_csv(arquivo)
tarefas = df.to_dict(orient="records")


relatorio = []



def executar_tarefa(tarefa):
    """Executa uma ação com base no tipo de tarefa e registra o resultado."""
    tipo = tarefa["Tipo"]
    dado = str(tarefa["Dado"])  
    inicio = time.time()

    try:
        if tipo == "abrir_navegador":
            pyautogui.press("win")

            pyautogui.write("Chrome", interval=0.2)
            webbrowser.open(dado)  
            time.sleep(5)  

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
            time.sleep(int(dado))  

        else:
            raise ValueError(f"Tipo desconhecido: {tipo}")

        status = "Sucesso"
    except Exception as e:
        status = f"Erro: {str(e)}"

    tempo_execucao = round(time.time() - inicio, 2)  
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
pyautogui.write("www.monstercat.com", interval=0.1)
time.sleep(2)
pyautogui.press("enter")
time.sleep(2)  

x_pagina, y_pagina = 379, 297  
pyautogui.click(x_pagina, y_pagina)
time.sleep(4)

x_menu, y_menu = 2189, 143  
pyautogui.click(x_menu, y_menu)
time.sleep(2)

x_about, y_about = 2215, 324  
pyautogui.click(x_about, y_about)
time.sleep(2)

x_about1, y_about1 = 2224, 361  
pyautogui.click(x_about1, y_about1)
time.sleep(1)
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

relatorio_arquivo = r"C:\Users\Pichau\Desktop\faculdade\automação\relatorio_execucao.xlsx"
os.makedirs(os.path.dirname(relatorio_arquivo), exist_ok=True)
wb.save(relatorio_arquivo)
print(f"Relatório salvo em: {relatorio_arquivo}")

