import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkBlue5')

EXCEL_FILE = 'Person_Data.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields : ')],
    [sg.Text('Name : ', size=(15, 1)), sg.InputText(key='Name', do_not_clear=True)],
    [sg.Text('Surname : ', size=(15,1)), sg.InputText(key='Surname', do_not_clear=True)],
    [sg.Text('Gender : ', size=(15,1)), sg.Combo(['Female','Male'], key='Gender', size=(7,1))],
    [sg.Text('City : ', size=(15,1)), sg.InputText(key='City', do_not_clear=True)],
    [sg.Text('Age : ', size=(15,1)), sg.Spin([i for i in range(1,100)], initial_value=0, key='Age', size=(3,1))],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Person Data Entry Form', layout)

def clear_input():
    for key in values:
        window[key]('')
        window['Name'].set_focus()

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index = False)
        sg.popup('Data saved!')
window.close()
