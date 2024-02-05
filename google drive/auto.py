import pyautogui

pyautogui.alert('Iniciando automação')
pyautogui.PAUSE = 0.7

# abrir o google drive
pyautogui.click('winleft')
pyautogui.write('opera')
pyautogui.click('enter')
pyautogui.write('https://drive.google.com/drive/home')
pyautogui.click('enter')

# voltar para a area de trabalho
pyautogui.hotkey('winleft', 'd')

# clickar no arquivo
pyautogui.moveTo('', '')
pyautogui.mouseDown()

# arrasta-lo ate um determinado local
pyautogui.moveTo('', '')

# voltar a aba do google drive
pyautogui.hotkey('alt', 'tab')

# soltar o arquivo
pyautogui.mouseUp()

