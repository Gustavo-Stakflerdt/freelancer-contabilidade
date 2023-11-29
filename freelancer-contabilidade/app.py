import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']  # Página de produtos

# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(910, 198, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(890, 292, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(876, 415, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(893, 502, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(891, 583, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(893, 666, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão de Próximo
    pyautogui.click(892, 724, duration=1)
    sleep(1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(890, 219, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(903, 303, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(880, 401, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(882, 472, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(888, 562, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(898, 595, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(894, 615, duration=1)
    else:
        pyautogui.click(877, 641, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(919, 652, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão de Próximo
    pyautogui.click(887, 715, duration=1)
    sleep(1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(903, 235, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(890, 313, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(866, 412, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(870, 533, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(876, 613, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão de Concluir
    pyautogui.click(865, 683, duration=1)
    sleep(1)

    # Botão de Ok
    pyautogui.click(1233, 170, duration=1)
    sleep(1)

    # Botão de Adicionar Mais Um
    pyautogui.click(1051, 455, duration=1)
    sleep(1)
