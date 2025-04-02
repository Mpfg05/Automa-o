import os
import csv
import pandas as pd
import pyautogui
import time
import webbrowser

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

            continuar = input("Deseja capturar mais posições? (s/n): ").strip().lower()
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

    #  Abrir YouTube primeiro
    webbrowser.open("https://www.youtube.com")
    time.sleep(5)  # Espera carregar

    # Lê as tarefas do CSV
    df = pd.read_csv("posicoes.csv")

    dados_relatorio= []

    for _, tarefa in df.iterrows():
        x, y, acao, texto = int(tarefa["x"]), int(tarefa["y"]), tarefa["acao"], tarefa["texto"]

        try:
            if acao == "clique":
                pyautogui.click(x, y)
                time.sleep(2)

            elif acao == "digitar":
                pyautogui.click(x, y)
                pyautogui.write(texto, interval=0.2)
                pyautogui.press("enter")
                time.sleep(3)

            print(f"Ação {acao} executada em ({x}, {y})!")

        except Exception as e:
            print(f"Erro ao executar {acao}: {e}")

executar_tarefas()
  
def gerar_relatorio:
  wb = Workbook()
  ws = wb.active
  ws.title = "Relatório de Execução de tarefas"
  ws.append(["Ação , "X", "Y", "Texto", "Status", "Tempo de Execução (s)"])
  for dados in dados: 
    ws.append(dados)
