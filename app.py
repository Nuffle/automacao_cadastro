from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from time import sleep
import openpyxl
import pyperclip
import pyautogui

navegador = webdriver.Chrome()
navegador.get("https://cadastro-produtos-devaprender.netlify.app/index.html")
sleep(5)

# Entrar na panilha.
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
0.94
# Copiar informações de um campo e colar no seu campo correspondente.
for linha in sheet_produtos.iter_rows(min_row=2):
   nome_produto = linha[0].value
   pyperclip.copy(nome_produto)
   navegador.find_element('xpath', '//*[@id="product_name"]').click()
   pyautogui.hotkey('ctrl','v')
   
   descricao = linha[1].value
   pyperclip.copy(descricao)
   navegador.find_element('xpath', '//*[@id="description"]').click()
   pyautogui.hotkey('ctrl','v')

   categoria = linha[2].value
   pyperclip.copy(categoria)
   navegador.find_element('xpath', '//*[@id="category"]').click()
   pyautogui.hotkey('ctrl','v')
   
   codigo_produto = linha[3].value
   pyperclip.copy(codigo_produto)
   navegador.find_element('xpath', '//*[@id="product_code"]').click()
   pyautogui.hotkey('ctrl','v')
   
   peso = linha[4].value
   pyperclip.copy(peso)
   navegador.find_element('xpath', '//*[@id="weight"]').click()
   pyautogui.hotkey('ctrl','v')
   
   dimensoes = linha[5].value
   pyperclip.copy(dimensoes)
   navegador.find_element('xpath', '//*[@id="dimensions"]').click()
   pyautogui.hotkey('ctrl','v')
   
   # Próxima página.
   navegador.find_element('xpath', '/html/body/div/form/div[7]/button[1]').click()
   sleep(1)
   
   preco = linha[6].value
   pyperclip.copy(preco)
   navegador.find_element('xpath', '//*[@id="price"]').click()
   pyautogui.hotkey('ctrl','v')
   
   qtd_estoque = linha[7].value
   pyperclip.copy(qtd_estoque)
   navegador.find_element('xpath', '//*[@id="stock"]').click()
   pyautogui.hotkey('ctrl','v')
   
   data_validade = linha[8].value
   pyperclip.copy(data_validade)
   navegador.find_element('xpath', '//*[@id="expiry_date"]').click()
   pyautogui.hotkey('ctrl','v')
   
   cor = linha[9].value
   pyperclip.copy(cor)
   navegador.find_element('xpath', '//*[@id="color"]').click()
   pyautogui.hotkey('ctrl','v')
   
   tamanho = linha[10].value
   navegador.find_element('xpath', '//*[@id="size"]').click()
   if tamanho == 'Pequeno':
      navegador.find_element('xpath', '//*[@id="size"]/option[1]').click()
   elif tamanho == 'Médio':
      navegador.find_element('xpath', '//*[@id="size"]/option[2]').click()
   else:
      navegador.find_element('xpath', '//*[@id="size"]/option[3]').click()
   
   material = linha[11].value
   pyperclip.copy(cor)
   navegador.find_element('xpath', '//*[@id="material"]').click()
   pyautogui.hotkey('ctrl','v')
   
   # Próxima página
   navegador.find_element('xpath', '/html/body/div/form/div[7]/button[1]').click()
   sleep(1)
   
   fabricante = linha[12].value
   pyperclip.copy(fabricante)
   navegador.find_element('xpath', '//*[@id="manufacturer"]').click()
   pyautogui.hotkey('ctrl','v')
   
   pais_origem = linha[13].value
   pyperclip.copy(pais_origem)
   navegador.find_element('xpath', '//*[@id="country"]').click()
   pyautogui.hotkey('ctrl','v')
   
   observacoes = linha[14].value
   pyperclip.copy(observacoes)
   navegador.find_element('xpath', '//*[@id="remarks"]').click()
   pyautogui.hotkey('ctrl','v')
   
   codigo_barras = linha[15].value
   pyperclip.copy(codigo_barras)
   navegador.find_element('xpath', '//*[@id="barcode"]').click()
   pyautogui.hotkey('ctrl','v')
   
   loc_armazem = linha[16].value
   pyperclip.copy(loc_armazem)
   navegador.find_element('xpath', '//*[@id="warehouse_location"]').click()
   pyautogui.hotkey('ctrl','v')
   
   # Botão concluir.
   navegador.find_element('xpath', '/html/body/div/form/div[6]/button[1]').click()
   sleep(1)
   
   # Modal OK.
   pyautogui.click(713,184,duration=0.5)
   
   # Adicionar mais.
   navegador.find_element('xpath', '/html/body/div/div/button').click()
   sleep(1)