import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1180,300, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1180,417, duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1180,560, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1180,677, duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1180,783, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1180,897, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1157,980, duration=1)
    

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1180,333, duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1180,428, duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1180,548, duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1180,643, duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1180,760, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1180,798, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1180,826, duration=1)
    else:
        pyautogui.click(1180,850, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1180,860, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1180,953, duration=1)


    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1180,353, duration=1)
    pyautogui.hotkey('ctrl','v')
                     
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1180,457, duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1180,575, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1180,735, duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1180,841, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1180,928, duration=1)
    pyautogui.click(1653,237, duration=2)
    pyautogui.click(1422,637, duration=2)
