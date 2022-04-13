import PySimpleGUI as sg
import pandas as pd

EXCEL_FILE = 'mastersheet.xlsx'
df = pd.read_excel(EXCEL_FILE)

sg.theme('Dark')
column1 = [[sg.Text('Total Materials', size=(10,1), justification='right'), sg.InputText(key='Material Cost')],
            [sg.Text('Total Labor', size=(10,1), justification='right'), sg.InputText(key='Labor Cost')],
            [sg.Text('Total Cost', size=(10,1), justification='right'), sg.InputText(key='Total Cost')],
            [sg.Text('Delivery Date', size=(10,1), justification='right'), sg.InputText(key='Delivery Date')]]
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Work Done For', size=(15,1), justification='left'), sg.InputText(key='Work Done For'),
        sg.Text('Order No.', size=(15,1), justification='right'), sg.InputText(key='Order No.')],
    [sg.Text('Professor', size=(15,1), justification='left'), sg.InputText(key='Professor'),
        sg.Text('Date', size=(15,1), justification='right'), sg.InputText(key='Date')],
    [sg.Text('Phone.', size=(15,1), justification='left'), sg.InputText(key='Phone'),
        sg.Text('Charge Account No.', size=(15,1), justification='right'), sg.InputText(key='Account')],
    [sg.Text('NetID', size=(15,1), justification='left'), sg.InputText(key='NetID'),
        sg.Text('Department', size=(15,1), justification='right'), sg.InputText(key='Department')],
    [sg.Column(column1, background_color='#F7F7F7', justification='left')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Invoice entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()