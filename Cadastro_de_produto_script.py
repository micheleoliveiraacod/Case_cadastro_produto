# Automação de cadastro de produtos em um sistema/site
# preparação do ambiente
import pandas
import openpyxl
import time
import pyautogui
import time
# pyautogui.click -> clica
# pyautogui.write -> escreve
# pyautogui.press -> aperta uma tecla
# pyautogui.hotkey -> aperta uma combinação de teclas
# pyautogui.sleep -> espera um tempo
# pyautogui.position -> pega a posição do mouse
# pyautogui.locateOnScreen -> localiza um elemento na tela
# pyautogui.center -> pega o centro de um elemento localizado na tela
# pyautogui.moveTo -> move o mouse para uma posição
# pyautogui.moveRel -> move o mouse em relação a posição atual
# pyautogui.scroll -> rola a tela para cima ou para baixo


pyautogui.PAUSE = 0.5  # pausa de meio segundo entre as ações para todos os comandos
Link = 'link'  # link do sistema/site

# abrir o navegador (chrome)
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

# entrar no link 
pyautogui.write("link")
pyautogui.press("enter")
time.sleep(3)

# Passo 2: Fazer login
# selecionar o campo de email
pyautogui.click(x=685, y=451)
# escrever o seu email
pyautogui.write("login")
pyautogui.press("tab") # passando pro próximo campo
pyautogui.write("senha")
pyautogui.click(x=955, y=638) # clique no botao de login
time.sleep(3)

# Passo 3: Importar a base de produtos pra cadastrar
import pandas as pd

tabela = pd.read_csv("produtos.csv")

print(tabela)

# Passo 4: Abrir a base de dados (importar o arquivo csv ou excel)
# se for excel e com abas, use este codigo:
# tabela = pandas.read_excel('produtos.xlsx', sheet_name='Planilha1')

tabela = pandas.read_csv('produtos.csv')  # substituir pelo nome do arquivo correto
for linha in tabela.index:
    produto = tabela.loc[linha, 'produto']
    preco = tabela.loc[linha, 'preco']
    quantidade = tabela.loc[linha, 'quantidade']
    
#Para cada linha dentro da minha tabela, executar o passo 4
for linha in tabela.index:
    # Passo 4: Cadastrar o produto
    pyautogui.click(x=500, y=300)  # clicar no campo do produto (substituir pelas coordenadas corretas)
    codigo = str(tabela.loc[linha, 'codigo'])
    pyautogui.write(codigo)
    pyautogui.press('tab')  # ir para o campo do preço

    marca = str(tabela.loc[linha, 'marca'])
    pyautogui.write(marca)
    pyautogui.press('tab')

    tipo = str(tabela.loc[linha, 'tipo'])
    pyautogui.write(tipo)
    pyautogui.press('tab')

    categoria = str(tabela.loc[linha, 'categoria'])
    pyautogui.write(categoria)
    pyautogui.press('tab')

    preco_unitario = str(tabela.loc[linha, 'preco_unitario'])
    pyautogui.write(preco_unitario)
    pyautogui.press('tab')  # ir para o campo da quantidade

    custo = str(tabela.loc[linha, 'quantidade_estoque'])
    pyautogui.write(custo)
    pyautogui.press('tab')

    margem = str(tabela.loc[linha, 'margem'])
    if margem != 'nan':  #se a observação for diferente de vazio, ele escreve a observação
    pyautogui.write(margem)
    pyautogui.press('tab')

    pyautogui.press("enter") # cadastra o produto (botao enviar)
    # dar scroll de tudo pra cima
    pyautogui.scroll(5000)
    # Passo 5: Repetir o processo de cadastro até o fim