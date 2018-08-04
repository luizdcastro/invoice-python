
import pyautogui
import time
import pandas as pd

df = pd.read_excel("postpaid.xlsx", "Sheet1")


def vat_number_invoice():
    time.sleep(3)
    pyautogui.click(x=885, y=400)
    vat = (df['CNPJ'][0])
    pyautogui.typewrite(str(vat))
    pyautogui.hotkey('enter')


def service_code():
    time.sleep(2)
    for __tab in range(8):
        pyautogui.hotkey('tab')
    pyautogui.hotkey('alt','down')
    for __down in range(8):
        pyautogui.hotkey('down')
    pyautogui.hotkey('enter')
    

def amount_invoice(): 
    time.sleep(3)
    pyautogui.click(x=893, y=581)
    pyautogui.typewrite(str('1'))
    pyautogui.hotkey('tab')
    time.sleep(1)
    value = (df['Total'][0])
    pyautogui.typewrite(str(value))
    pyautogui.hotkey('tab') 


def text_body_invoice():
    time.sleep(2)
    description = (df['Descrição'][0])
    pyautogui.typewrite(str(description))


def due_date_invoice():    
    time.sleep(1)
    due_date = (df['Vencimento'][0])
    br_format_date = due_date.strftime('%d/%m/%Y')
    pyautogui.typewrite(str(br_format_date))


def confirm_invoice():
    time.sleep(1)
    for __tab in range(8):
        pyautogui.hotkey('tab')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(4)
    pyautogui.click(x=890, y=930) 
    time.sleep(4)
    pyautogui.click(x=1135, y=393) 


def save_invoice():
    time.sleep(4)
    pyautogui.hotkey('ctrl', 's')
    time.sleep(2)
    save_pk = (df['PK'][0])
    pyautogui.typewrite(str(save_pk))
    time.sleep(1)
    save_offer = (df['Offer'][0])
    pyautogui.typewrite(str(save_offer))
    time.sleep(1)
    pyautogui.hotkey('tab', 'tab', 'enter', interval=0.1)
    pyautogui.hotkey('ctrl', 'w')


def next_index():
    time.sleep(1)
    df.drop(0, inplace=True)
    df.reset_index(drop=True, inplace=True)


def new_issue():
    time.sleep(3)
    pyautogui.click(x=704, y=229)
    time.sleep(4)
    i = 0 
    while i < 100: 
        issue_invoice_baruei()
        if i == 100:
            break
        i += 1


def issue_invoice_baruei():
          
    vat_number_invoice()      
    service_code()
    amount_invoice()
    text_body_invoice() #class for body
    due_date_invoice()  #class for body      
    confirm_invoice()   
    save_invoice()   
    next_index()       
    new_issue()

issue_invoice_baruei()


            
