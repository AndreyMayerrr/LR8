from openpyxl import load_workbook
import PySimpleGUI as sg
from datetime import datetime

element_size = (25, 1)

layout = [
    [sg.Text('Фамилия', size=element_size), sg.Input(key='familia', size=element_size)],
    [sg.Text('Имя', size=element_size), sg.Input(key='name', size=element_size)],
    [sg.Text('Отчество', size=element_size), sg.Input(key='otchestvo', size=element_size)],
    [sg.Text('Номер страхового полиса', size=element_size), sg.Input(key='nomer', size=element_size)],
    [sg.Text('Номер паспорта', size=element_size), sg.Input(key='nomerp', size=element_size)],
    [sg.Text('Серия паспорта', size=element_size), sg.Input(key='seriap', size=element_size)],
    [sg.Text('Дата рождения', size=element_size), sg.Input(key='data1', size=element_size)],
    [sg.Text('Пол', size=element_size), sg.Input(key='pol', size=element_size)],
    [sg.Text('Дата обращения', size=element_size), sg.Input(key='data2', size=element_size)],
    [sg.Text('Симптомы и жалобы', size=element_size), sg.Input(key='sim', size=element_size)],
    [sg.Text('Дата выписки', size=element_size), sg.Input(key='data3', size=element_size)],
    [sg.Text('Состояние здоровья', size=element_size), sg.Input(key='zdor', size=element_size)],
    [sg.Button('Добавить'), sg.Button('Закрыть')]
]

window = sg.Window('Учет больных в стационаре', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':
        break
    if event == 'Добавить':
        try:
            wb = load_workbook('больница.xlsx')
            sheet = wb['Лист1']
            ID = len(sheet['ID'])
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [
                ID,
                values['familia'],
                values['name'],
                values['otchestvo'],
                values['nomer'],
                values['nomerp'],
                values['seriap'],
                values['data1'],
                values['pol'],
                values['data2'],
                values['sim'],
                values['data3'],
                values['zdor'],
                time_stamp
            ]
            sheet.append(data)
            wb.save('больница.xlsx')

            # Очистка полей ввода
            for key in values:
                window[key].update(value='')
            window['name'].set_focus()
            sg.popup('Данные сохранены')
        except PermissionError:
            sg.popup('Ошибка доступа', 'Файл используется другим пользователем.\nПопробуйте позже.')


window.close()