from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

sg.theme('darkgrey')

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Patient Firstname'),sg.Push(),sg.Input(key='PATIENT_FIRST_NAME')],
    [sg.Text('Patient Lastname'),sg.Push(),sg.Input(key='LAST_NAME')],
    [sg.Text('Mobile Number'),sg.Push(),sg.Input(key='MOBILE_NUMBER')],
    [sg.Text('Admit Time'),sg.Push(),sg.Input(key='ADMIT_TIME')],
    [sg.Text('Admit Cause'),sg.Push(),sg.Input(key='ADMIT_CAUSE')],
    [sg.Text('Doctor Reffered'),sg.Push(),sg.Input(key='DOCTOR_REFFERED')],
    [sg.Button('Submit'),sg.Button('Close')]
]
window =sg.Window('Data Entry',layout,element_justification='centre')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Close':
        break
    if event =='Submit':
        try:
            wb = load_workbook('data_entry.xlsx')
            sheet = wb['Sheet1'] 

            data = [values['PATIENT_FIRST_NAME'],values['LAST_NAME'],values['MOBILE_NUMBER'],values['ADMIT_TIME'],values['ADMIT_CAUSE'],values['DOCTOR_REFFERED']]
            sheet.append(data)     
            wb.save('data_entry.xlsx')  

            window['PATIENT_FIRST_NAME'].update(value='')
            window['LAST_NAME'].update(value='')
            window['MOBILE_NUMBER'].update(value='')
            window['ADMIT_TIME'].update(value='')
            window['ADMIT_CAUSE'].update(value='')
            window['DOCTOR_REFFERED'].update(value='')

            window['PATIENT_FIRST_NAME'].set_focus()

            sg.popup('Success','Data Saved')
        except PermissionError:
          sg.popup('File in use','File is being used by another User. \nPlease try again later.')

          window.close()
