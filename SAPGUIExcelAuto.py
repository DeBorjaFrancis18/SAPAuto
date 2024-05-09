import openpyxl
from datetime import datetime, timedelta
from pyrfc import Connection
import pyautogui
import pytesseract
import time
import pyperclip


delay = 5
short = 3

def create_sales_order():
    
    time.sleep(delay)

    #click on the bar
    pyautogui.click(x=115, y=105)

    #t-code
    pyautogui.write('/nVA01')
    pyautogui.press('enter')

    time.sleep(delay)

    pyautogui.write(header_info['OrderType']) 
    pyautogui.press('tab')
    pyautogui.write(header_info['SalesOrg']) 
    pyautogui.press('tab')
    pyautogui.write(header_info['Division'])
    pyautogui.press('tab')
    pyautogui.write(header_info['DistributionChannel'])
    pyautogui.press('tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.press('end')
    pyautogui.press('enter')

    time.sleep(delay)

    #Sales Order info
    pyautogui.write(str(header_info['SoldToParty']))
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(header_info['ShipToParty'])  
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(header_info['PONumber'])
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(header_info['PODate'])
    pyautogui.press('tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.press('end')
    pyautogui.press('enter')
    time.sleep(short)
    pyautogui.press('enter')

    time.sleep(delay)

    #shipping info
    pyautogui.click(x=220, y=345)
    time.sleep(delay)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(header_info['DeliveryDate'])
    time.sleep(short)
    pyautogui.press('tab')
    time.sleep(short)
    pyautogui.write(str(header_info['DeliveryPlant']))  
    pyautogui.press('tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.press('end')
    pyautogui.press('enter')
    time.sleep(short)
    pyautogui.press('enter')

    time.sleep(delay)

    print('PO Date: ' + header_info['PODate'])
    print('PO Number: '+ header_info['PONumber'])
    print('RDD: ' + header_info['DeliveryDate'])

    #materials info
    
    for item in item_info:
        print(item['Material'], type(item['Material']))
        pyautogui.write(str(item['Material']))
        pyautogui.press('tab')
        pyautogui.write(str(item['Qty']))  # Convert to string if necessary
        pyautogui.press('tab')
        pyautogui.write(item['UoM'])
        pyautogui.press('tab')
        pyautogui.press('enter')
        time.sleep(delay)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(delay)

    pyautogui.click(x=210, y=1010)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    Status = pyperclip.paste()
    print(Status)

# Read Excel file
wb = openpyxl.load_workbook('SAPAuto.xlsx')
header_sheet = wb['Header']
item_sheet = wb['Item']

# Extract header information
header_info = {
    'OrderType': header_sheet['A2'].value,
    'SalesOrg': header_sheet['B2'].value,
    'DistributionChannel': header_sheet['C2'].value,
    'Division': header_sheet['D2'].value,
    'SoldToParty': int(header_sheet['E2'].value),
    'ShipToParty': header_sheet['F2'].value,
    'DeliveryPlant': header_sheet['G2'].value,
    'PONumber': 'TA_' + datetime.now().strftime("%Y%m%d%H%M%S"),  # Generating PO Number
    'PODate': (datetime.now() - timedelta(days=1)).strftime("%d.%m.%y"),  # Generating PO Date
    'DeliveryDate': (datetime.now() + timedelta(days=header_sheet['H2'].value)).strftime("%d.%m.%y"),
    'ShippingCondition': header_sheet['I2'].value,
    'Vendor': header_sheet['J2'].value,
}

# Extract item information
item_info = []
for row in item_sheet.iter_rows(min_row=2, values_only=True):
    item_info.append({
        'Material': row[0],
        'Qty': row[1],
        'UoM': row[2],
        'Sloc': row[3]
    })

create_sales_order()
