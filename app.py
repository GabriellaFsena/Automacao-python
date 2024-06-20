import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1325,324,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1321,420,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1331,566,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1324,667,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1322,768,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1323,860,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Botão próximo
    pyautogui.click(1345,931,duration=1)
    sleep(3)

    # # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1327,350,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1334,449,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1329,544,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1327,641,duration=1)
    pyautogui.hotkey('ctrl','v')

    # # Tamanho
    tamanho = linha[10].value
    pyautogui.click(1350,742,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1348,773,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1350,794,duration=1)
    else:
        pyautogui.click(1349,821,duration=1)

    # # material    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1324,833,duration=1)
    pyautogui.hotkey('ctrl','v')
    #Botão próximo
    pyautogui.click(1334,899,duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1319,374,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1319,470,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1322,560,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1323,714,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1326,819,duration=1)
    pyautogui.hotkey('ctrl','v')

     # Botão concluir
    pyautogui.click(1335,884,duration=1)
    # Botão confirmar inclusão
    pyautogui.click(1806,231,duration=1)
    #  Botão confirmação 2
    pyautogui.click(1603,622,duration=1)
    #  # iniciar cadastro novamente
    # pyautogui.click(1695,580,duration=1)