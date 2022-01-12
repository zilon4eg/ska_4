import os
import PySimpleGUI
import registry


class GUI:
    def __init__(self):
        """
        Получить текущую директорию, где запущен скрипт
        dir = os.path.abspath(os.curdir)
        Получить текущую директорию, где расположен скрипт
        os.path.abspath(__file__)
        """
        position = os.path.abspath(__file__).rfind('\\')
        self.config_path = f'{os.path.abspath(__file__)[:position]}\\config.ini'

    def load_settings(self):
        settings = dict()
        with open(self.config_path, 'r', encoding='cp1251') as file:
            for line in file:
                line = line.strip().split('=')
                settings.update({line[0]: line[1]})
        return settings

    def save_settings(self, new_settings):
        settings = self.load_settings()
        for setting in new_settings:
            settings.update({setting: new_settings[setting]})
            # print(settings)
        with open(self.config_path, 'w', encoding='cp1251') as file:
            for line in settings:
                file.writelines(f'{line}={settings[line]}\n')

    def main_menu(self):
        settings = self.load_settings()

        # ------ Menu Definition ------ #
        menu_def = [
            ['File', ['Exit']],
            ['Settings', ['Font', 'Hyperlink']],
            # ['Help', 'About'],
        ]
        # ----------------------------- #
        layout = [
            [
                PySimpleGUI.Menu(menu_def, tearoff=False)
            ],
            [
                PySimpleGUI.Text('Путь к файлу реестра: ', size=(17, 1)),
                PySimpleGUI.InputText(key='file', size=(58, 1)),
                PySimpleGUI.FileBrowse(target='file', initial_folder=settings['base_registry_path'], size=(7, 1))
            ],
            [
                PySimpleGUI.Text('Название листа в книге Excel: ', size=(23, 1)),
                PySimpleGUI.InputText(key='SHEET', size=(32, 1), disabled=True),
                PySimpleGUI.Checkbox('Использовать активный лист', default=True, key='WS_CHECKBOX', enable_events=True)
            ],
            [
                PySimpleGUI.Text('Путь к папке со сканами: ', size=(19, 1)),
                PySimpleGUI.InputText(key='folder', size=(56, 1)),
                PySimpleGUI.FolderBrowse(target='folder', initial_folder=settings['base_scan_path'], size=(7, 1))
            ],
            [
                PySimpleGUI.Output(size=(88, 20))
            ],
            [
                PySimpleGUI.Submit(button_text='Start'),
                PySimpleGUI.Cancel(button_text='Exit'),
            ]
        ]

        window_main = PySimpleGUI.Window('Hyperlinks creator v2.4.3', layout)

        while True:  # The Event Loop
            settings = self.load_settings()
            event, values = window_main.read()
            # print(event, values) #debug

            if values['WS_CHECKBOX'] is True:
                window_main['SHEET'].update('', disabled=True)
            else:
                window_main['SHEET'].update(disabled=False)

            if event in (None, 'Exit', 'Cancel'):
                break

            elif event in 'Font':
                window_main.hide()
                self.font_menu(settings)
                window_main.UnHide()

            elif event in 'Hyperlink':
                window_main.hide()
                self.color_chooser_menu(settings)
                window_main.UnHide()

            elif event in 'Start':
                base_registry_path = values['file']  # r'\\fs\SHARE\Documents\OTDEL-SECRETARY\Регистрация документов\Реестры'
                base_scan_path = values['folder']  # r'\\fs\SHARE\Documents\OTDEL-SECRETARY\Регистрация документов'
                if base_registry_path and base_scan_path:
                    base_registry_path = base_registry_path[:base_registry_path.rfind('/')]
                    self.save_settings({'base_registry_path': base_registry_path, 'base_scan_path': base_scan_path})

                registry_path = values['file']

                if not values['WS_CHECKBOX']:
                    ws_name = values['SHEET']
                else:
                    ws_name = True

                dir_scan = values['folder']
                if os.path.exists(dir_scan):
                    print(f'Доступность каталога сканов проверена.')
                else:
                    print(f'Каталог сканов недоступен.')
                    break

                registry.body(registry_path, dir_scan, ws_name, settings)

    def font_menu(self, settings):
        # settings = self.load_settings()
        font_list = settings['font_list'].split(',')
        font_name = settings['font_name']
        font_size = int(settings['font_size'])

        layout = [
            [PySimpleGUI.Combo(font_list, default_value=font_name, key='drop-down', enable_events=True)],
            [PySimpleGUI.Spin([sz for sz in range(10, 21)], font='Arial 20', initial_value=font_size,
                              change_submits=True,
                              key='spin'),
             PySimpleGUI.Slider(range=(10, 20), orientation='h', size=(10, 25),
                                change_submits=True, key='slider', font=f'{font_name.replace(" ", "")} 20', default_value=font_size),
             PySimpleGUI.Text("Ab", size=(2, 1), font=f'{font_name.replace(" ", "")} {str(font_size)}', key='text')],
            [
                PySimpleGUI.Submit(button_text='Ok'),
                PySimpleGUI.Cancel(button_text='Cancel')
            ]
        ]

        # sz = font_size
        window = PySimpleGUI.Window("Font size selector", layout, grab_anywhere=False)
        # Event Loop

        while True:
            event, values = window.read()

            window['text'].update(font=f'{values["drop-down"].replace(" ", "")} {str(font_size)}')

            if event in (PySimpleGUI.WIN_CLOSED, 'Cancel'):
                window.close()
                break
            sz_spin = int(values['spin'])
            sz_slider = int(values['slider'])
            sz = sz_spin if sz_spin != font_size else sz_slider
            if sz != font_size:
                font_size = sz
                font = f'{values["drop-down"].replace(" ", "")} {str(font_size)}'
                window['text'].update(font=font)
                window['slider'].update(sz)
                window['spin'].update(sz)

            if event in 'Ok':
                self.save_settings({'font_size': font_size, 'font_name': values['drop-down']})
                window.close()
                break

    def color_chooser_menu(self, settings):
        img_color = settings['hyperlink_color']

        layout = [
            [
                PySimpleGUI.Text('Код цвета:'),
                PySimpleGUI.Input(key='COLOR', readonly=True, size=(7, 1), enable_events=True),
                PySimpleGUI.ColorChooserButton(button_text='Choose color', key='COLOR')
            ],
            [
                PySimpleGUI.Submit(button_text='Ok'),
                PySimpleGUI.Cancel(button_text='Cancel'),
                PySimpleGUI.Text(' Пример цвета:'),
                PySimpleGUI.Button(button_text='', button_color=img_color, size=(2, 1), disabled=True, key='IMG_COLOR'),
            ],
        ]

        window = PySimpleGUI.Window("Font size selector", layout, grab_anywhere=False)
        # Event Loop

        while True:
            event, values = window.read()

            if values['COLOR'] in [None, 'None', '']:
                img_color = '#0563c1'
            else:
                img_color = values['COLOR']
            window['IMG_COLOR'].update(button_color=img_color)

            if event in (PySimpleGUI.WIN_CLOSED, 'Cancel'):
                window.close()
                break

            if event in 'Ok':
                hyperlink_color = values['COLOR']
                if hyperlink_color in [None, 'None', '']:
                    hyperlink_color = '#0563c1'
                self.save_settings({'hyperlink_color': hyperlink_color})
                window.close()
                break

    @staticmethod
    def progress_bar(size):
        # layout the window
        layout = [
            [PySimpleGUI.Text('Working...')],
            [PySimpleGUI.ProgressBar(size, orientation='h', size=(28, 20), key='PROGRESSBAR')]
        ]
        return PySimpleGUI.Window('Create hyperlinks', layout, disable_minimize=True, keep_on_top=True)


if __name__ == '__main__':
    gg = GUI()
    # gg.progress_bar_menu(50)
    gg.main_menu()
    pass
