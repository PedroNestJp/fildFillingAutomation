import openpyxl 
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
        # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(877,248, duration=0.6)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')

        # inicar processo em nova pagina
    pyautogui.click(1231,165,duration=0.6)
    sleep(2)
    pyautogui.hotkey('tab')

        # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    
        # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')

        # Tamanho
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
    elif tamanho == 'Médio':
            pyautogui.hotkey('down')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
    elif tamanho == 'Grande':
            pyautogui.hotkey('down')
            pyautogui.hotkey('down')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
    else : 
            pyautogui.hotkey('tab')
           
        
        # material 
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')

        # botão Próximo
    pyautogui.hotkey('enter')
    sleep(3)
    pyautogui.click(1048,218, duration=0.6)
    pyautogui.hotkey('tab')

        # fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
        # pais de origin
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
        # observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
        # código de barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
        # armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('enter')
    sleep(2.5)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')