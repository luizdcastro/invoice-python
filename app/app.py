
import pyautogui
import time
import pandas as pd

# Importing our database with all offers to issue invoices.
df = pd.read_excel("postpaid.xlsx", "Sheet1")

# Get and fill vat number from database.
def vat_number_invoice():
    time.sleep(3)
    pyautogui.click(x=885, y=400)
    vat = (df['CNPJ'][0])
    pyautogui.typewrite(str(vat))
    pyautogui.hotkey('enter')

# Select the service code and fill the field.
def service_code():
    time.sleep(3)
    for __tab in range(8):
        pyautogui.hotkey('tab')
    pyautogui.hotkey('alt','down')
    for __down in range(8):
        pyautogui.hotkey('down')
    pyautogui.hotkey('enter')
    
# Get the amount of the invoice and fill the field.
def amount_invoice(): 
    time.sleep(2)
    pyautogui.click(x=893, y=581)
    pyautogui.typewrite(str('1'))
    pyautogui.hotkey('tab')
    time.sleep(1)
    value = (df['Total'][0])
    pyautogui.typewrite(str(value))
    pyautogui.hotkey('tab') 

# Get the invoice description.
def text_body_invoice():
    time.sleep(2)
    description = (df['Description'][0])
    pyautogui.typewrite(str(description))

# Var that will fill the due date on invoice description.
def due_date_invoice():    
    time.sleep(1)
    due_date = (df['Due date'][0])
    br_format_date = due_date.strftime('%d/%m/%Y')
    pyautogui.typewrite(str(br_format_date))

# After the all data were inserted confirm the invoice issuance in two steps.
def confirm_invoice():
    time.sleep(1)
    for __tab in range(8):
        pyautogui.hotkey('tab')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(5)
    pyautogui.click(x=890, y=930) 
    time.sleep(5)
    pyautogui.click(x=1135, y=393) 

# Save the invoice in PDF format named with the offer number and close the webpage.
def save_invoice():
    time.sleep(5)
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)
    save_pk = (df['PK'][0])
    pyautogui.typewrite(str(save_pk))
    time.sleep(1)
    save_offer = (df['Offer'][0])
    pyautogui.typewrite(str(save_offer))
    time.sleep(2)
    pyautogui.click(x=1221, y=753) 
    pyautogui.hotkey('tab', 'tab', 'enter', interval=0.1)
    time.sleep(3)
    pyautogui.hotkey('ctrl', 'w')

# Delete the previous offer and go to the next.
def next_index():                               
    time.sleep(2)
    df.drop(0, inplace=True)
    df.reset_index(drop=True, inplace=True)

# Set the state to issue the next invoice.
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

# Agroup all the previous steps in a single function to issue invoices and save in sequence.
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
    
# Grab a coffe and watch the magic xD
issue_invoice_baruei()



