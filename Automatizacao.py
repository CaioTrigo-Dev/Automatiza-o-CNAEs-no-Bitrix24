import pyautogui
import time
import pandas as pd
import pyperclip

caminho = r'C:\Users\caio_\Downloads\adicionar.png'
url = 'https://promobrindes.bitrix24.com.br/page/inteligencia_dados/fornecedores/type/1320/details/0/?categoryId=280&st%5Btool%5D=crm&st%5Bc_section%5D=custom_section&st%5Bc_sub_section%5D=kanban&st%5Bc_element%5D=create_button&st%5Bp1%5D=crmMode_classic&st%5Bcategory%5D=entity_operations&st%5Bevent%5D=entity_add_open&st%5Btype%5D=dynamic'

df = pd.read_excel(r'C:\Users\caio_\Downloads\Tabela.xlsx')
pyautogui.PAUSE = 0.5

pyautogui.press('win')
pyautogui.write('google chrome')
pyautogui.press('enter')
pyautogui.write(url)
pyautogui.press('enter')
time.sleep(3)
pyautogui.hotkey('PgDn')
pyautogui.click(x=266, y=697)
pyautogui.click(x=420, y=276)
pyautogui.click(x=309, y=479)
pyautogui.hotkey('ctrl', 'a')
pyautogui.write('Atividade Economica')
pyautogui.press('tab')
pyautogui.write('Teste')
# pyautogui.click(x=144, y=607)

try:
    
    x, y = pyautogui.locateCenterOnScreen(caminho)
    
    
    for i in range(len(df)):
        x, y = pyautogui.locateCenterOnScreen(caminho)
        pyautogui.click(x, y)
        cnae = df.loc[i, 'CNAE']
        descricao = df.loc[i, 'DESCRIÇÃO']
        pyperclip.copy(f'{cnae} - {descricao}')
        # pyautogui.write(f'{cnae} - {descricao}')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.scroll(-250)
        # pyautogui.move(x=10, y=100)
except pyautogui.ImageNotFoundException:
    print("Botão 'Adicionar' não encontrado na tela")
    
except Exception as e:
    print(f"Erro: {e}")