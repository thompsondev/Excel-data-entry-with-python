from asyncio import events
import imp
from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

sg.theme('DarkAmber')

layout =  [[sg.Text('LOCATION' ),sg.Push(), sg.Input(key='LOCATION')],
           [sg.Text('OPERATIVE NAME'),sg.Push(), sg.Input(key='OPERATIVE_NAME')],
           [sg.Text('OPERATIVE ADDRESS'),sg.Push(), sg.Input(key='OPERATIVE_ADDRESS')],
           [sg.Text('MOBILE NUMBER'),sg.Push(), sg.Input(key='MOBILE_NUMBER')],
           [sg.Text('GUARANTOR 1'),sg.Push(), sg.Input(key='GUARANTOR_1')],
           [sg.Text('GUARANTOR 1 ADDRESS'),sg.Push(), sg.Input(key='GUARANTOR_1_ADDRESS')],
           [sg.Text('GUARANTOR 1 MOBILE'),sg.Push(), sg.Input(key='GUARANTOR_1_MOBILE')],
           [sg.Text('GUARANTOR 2'),sg.Push(), sg.Input(key='GUARANTOR_2')],
           [sg.Text('GUARANTOR 2 ADDRESS'),sg.Push(), sg.Input(key='GUARANTOR_2_ADDRESS')],
           [sg.Text('GUARANTOR 2 MOBILE'),sg.Push(), sg.Input(key='GUARANTOR_2_MOBILE')],
           [sg.Text('REMARKS'),sg.Push(), sg.Input(key='REMARKS')],
           [sg.Button('Submit'), sg.Button('Close')]]

Window = sg.Window('RH PRIVATE SECURITY OPERATIVE DATA', layout,element_justification='center')

while True:
    event, values = Window.read()
    if event == sg.WIN_CLOSED or event == 'Close':
        break
    if event == 'Submit':
        try:
            wb = load_workbook('Book1.xlsx')
            sheet = wb['Sheet1']
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [ID, values['LOCATION'], values['OPERATIVE_NAME'], values['OPERATIVE_ADDRESS'], values['MOBILE_NUMBER'], values['GUARANTOR_1'], values['GUARANTOR_1_ADDRESS'], values['GUARANTOR_1_MOBILE'], values['GUARANTOR_2'], values['GUARANTOR_2_ADDRESS'], values['GUARANTOR_2_MOBILE'], values['REMARKS'], time_stamp]

            sheet.append(data)

            wb.save('Book1.xlsx')

            Window['LOCATION'].update(value='')
            Window['OPERATIVE_NAME'].update(value='')
            Window['OPERATIVE_ADDRESS'].update(value='')
            Window['MOBILE_NUMBER'].update(value='')
            Window['GUARANTOR_1'].update(value='')
            Window['GUARANTOR_1_ADDRESS'].update(value='')
            Window['GUARANTOR_1_MOBILE'].update(value='')
            Window['GUARANTOR_2'].update(value='')
            Window['GUARANTOR_2_ADDRESS'].update(value='')
            Window['GUARANTOR_2_MOBILE'].update(value='')
            Window['REMARKS'].update(value='')
            Window['LOCATION'].set_focus()

            sg.popup('Success', 'Data Saved')
        except PermissionError:
            sg.popup('File in use', 'File is being used by another User.\n Please try again later')




Window.close()
           