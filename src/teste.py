import os 

diretorio_atual = os.getcwd()
caminho_planilha = os.path.join(diretorio_atual, 'assets', 'processos.xlsx')

print(caminho_planilha)