import os
import PySimpleGUI as sg
import xls


def light_files_in_dir(list_files):
    list_files = list(file for file in list_files if not file[:file.rfind('.')].isdigit())
    list_files = list(map(lambda x: x[x.rfind('№') + 1:x.rfind('.')].lower(), list_files))
    return list_files


def miss_files(list1, list2):
    # miss_list = list(file for file in list1 if file not in list2)
    miss_list = list(set(list1) - set(list2))
    return miss_list


if __name__ == '__main__':
    base_registry_path = r'\\fs\SHARE\Documents\OTDEL-SECRETARY\Регистрация документов\Реестры'
    base_scan_path = r'\\fs\SHARE\Documents\OTDEL-SECRETARY\Регистрация документов'

    layout = [
        [
            sg.Text('Путь к файлу реестра: ', size=(17, 1)),
            sg.InputText(key='file', size=(58, 1)),
            sg.FileBrowse(target='file',
                          initial_folder=base_registry_path, size=(7, 1))
        ],
        [
            sg.Text('Название листа в книге Excel: ', size=(23, 1)),
            sg.InputText(key='sheet', size=(62, 1))
        ],
        [
            sg.Text('Путь к папке со сканами: ', size=(19, 1)),
            sg.InputText(key='folder', size=(56, 1)),
            sg.FolderBrowse(target='folder',
                            initial_folder=base_scan_path, size=(7, 1))
        ],
        [
            sg.Output(size=(88, 20))
        ],
        [
            sg.Submit(button_text='Start'),
            sg.Cancel(button_text='Exit')
        ]
    ]

    window = sg.Window('Hyperlinks creator', layout)
    while True:                             # The Event Loop
        event, values = window.read()
        # print(event, values) #debug

        if event in (None, 'Exit', 'Cancel'):
            break

        if event == 'Start':
            registry_path = values['file']
            ws_name = values['sheet']
            dir_scan = values['folder']
            # registry_path = r'C:\Users\suhorukov.iv\Desktop\Реестр входящих 2020-2021.xlsx'
            # ws_name = r'2021'
            # dir_scan = r'\\fs\SHARE\Documents\OTDEL-SECRETARY\Регистрация документов\ВХОДЯЩИЕ 2021'

            if os.path.exists(registry_path) and os.path.exists(dir_scan):
                print(f'Доступность путей проверена.')
            else:
                print(f'Один из путей недоступен или не существует.')
                break

            xxl = xls.Excel(registry_path, dir_scan, ws_name)

            if registry_path[registry_path.rfind('\\') + 1:registry_path.rfind('.')].lower().count('исходящ') > 0:
                file_pref = 'исх.№'
                print('Загруженный документ идентифицирован как реестр Исходящих.')
            elif registry_path[registry_path.rfind('\\') + 1:registry_path.rfind('.')].lower().count('входящ') > 0:
                file_pref = 'вход.№'
                print('Загруженный документ идентифицирован как реестр Входящих.')
            else:
                file_pref = '№'
                print('Реестр не идентифицирован.')

            files_a = xxl.get_column()
            files_a_sort = list(map(lambda x: x.replace(r'/', r'-').strip().split()[0], files_a))
            files_dir = os.listdir(path=dir_scan)
            print(f'Получен список файов в папке {dir_scan}.')

            files_dir_clear = light_files_in_dir(files_dir)
            miss_list = miss_files(files_a_sort, files_dir_clear)
            miss_list.sort()
            for miss in miss_list:
                print(f'Не найдено сопоставление регистрационному номеру {miss} среди файлов.')

            print('Формирование гиперссылок.')
            for position, file_a in enumerate(files_a, 3):
                file_a_clear = file_a.replace(r'/', r'-').strip().split()[0]

                for file_dir in files_dir:
                    if not file_dir.isdigit():
                        file_type = file_dir[file_dir.rfind('.'):].lower()
                        file_dir_clear = file_dir[file_dir.rfind('№') + 1:file_dir.rfind('.')].lower()

                        if file_dir_clear == file_a_clear:
                            name = f'{file_pref.capitalize()}{file_a_clear}{file_type}'
                            link_name = f'{file_pref.upper()}{file_a_clear}{file_type}'
                            xxl.create_hyperlinks(name, link_name, position)
                            # print(position, name, link_name)
            print('Гиперссылки сформированы.')
            xxl.save()
            print(f'Файл {registry_path} сохранен')
            print('Complete...' + '\n' * 1)
