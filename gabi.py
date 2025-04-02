import os
import csv
import pandas as pd
import pyautogui
import time
import webbrowser
from openpyxl import Workbook

#  Função para capturar posições do mouse
def capturar_posicoes(nome_arquivo="posicoes.csv"):
    print("Para capturar a posição do mouse, arraste para o lugar desejado e aguarde!")
    dados = []  # Lista para armazenar coordenadas e ações

    while True:
        try:
            input("Posicione o mouse e pressione Enter... (ou digite 'sair' para finalizar) ")
            x, y = pyautogui.position()
            print(f"Posição capturada: X={x}, Y={y}")

            acao = input("Digite a ação (clique/digitar): ").strip().lower()
            texto = input("Se for digitar, escreva o texto (ou deixe vazio): ").strip()

            dados.append([x, y, acao, texto])
            
            if continuar != 's':
                break

        except KeyboardInterrupt:
            print("\nCaptura de posições interrompida!")
            break

    #  Salva os dados no arquivo CSV
    with open(nome_arquivo, "w", newline="") as b:
        writer = csv.writer(b)
        writer.writerow(["x", "y", "acao", "texto"])  # Cabeçalho
        writer.writerows(dados)

    print(f"Posições salvas em '{nome_arquivo}'!")

capturar_posicoes()

#  Função para ler a tabela e executar ações automaticamente
def executar_tarefas():
    webbrowser.open("https://www.youtube.com")
    time.sleep(5)
    inicio = time.time()

    df = pd.read_csv("posicoes.csv")
    dados_relatorio = []

    for _, tarefa in df.iterrows():
        x, y, acao, texto = int(tarefa["x"]), int(tarefa["y"]), tarefa["acao"], tarefa["texto"]

        try:
            inicio = time.time()
            if acao == "clique":
                pyautogui.click(x, y)
                time.sleep(2)
            elif acao == "digitar":
                pyautogui.click(x, y)
                pyautogui.write(texto, interval=0.2)
                pyautogui.press("enter")
                time.sleep(3)

            fim = time.time()
            tempo_execucao = round(fim - inicio, 2)
            print(f"Ação {acao} executada em ({x}, {y})!")

            dados_relatorio.append([acao, x, y, texto, "Sucesso", tempo_execucao])

        except Exception as e:
            print(f"Erro ao executar {acao}: {e}")
            dados_relatorio.append([acao, x, y, texto, f"Erro: {e}", 0])

    gerar_relatorio(dados_relatorio)

def gerar_relatorio(dados_relatorio):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório de Execução de Tarefas"

    ws.append(["Ação", "X", "Y", "Texto", "Status", "Tempo de Execução (s)"])
    
    for linha in dados_relatorio:
        ws.append(linha)

    wb.save("relatorio_execucao.xlsx")
    print("Relatório gerado com sucesso: 'relatorio_execucao.xlsx'")

executar_tarefas()
