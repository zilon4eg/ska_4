import os
import PySimpleGUI


class GUI:
    def __init__(self):
        self.config_path = f'{os.path.abspath(os.curdir)}\\config.ini'

    def load_settings(self):
        settings = dict()
        with open(self.config_path, 'r') as file:
            for line in file:
                line = line.strip().split('=')
                settings.update({line[0]: line[1]})
        return settings

    def save_settings(self, new_settings):
        settings = self.load_settings()
        for setting in new_settings:
            settings.update({setting: new_settings[setting]})
            print(settings)
        with open(self.config_path, 'w') as file:
            for line in settings:
                file.writelines(f'{line}={settings[line]}\n')

    def main_menu(self):
        # ------ Menu Definition ------ #
        menu_def = [
            ['File', ['Exit']],
            ['Settings', ['Font', 'Font size', 'Hyperlink color']],
            ['Help', 'About'],
        ]
        # ----------------------------- #
        layout = [
            [
                PySimpleGUI.Menu(menu_def, tearoff=False)
            ],
            [
                PySimpleGUI.Text('Путь к файлу реестра: ', size=(17, 1)),
                PySimpleGUI.InputText(key='file', size=(58, 1)),
                PySimpleGUI.FileBrowse(target='file', size=(7, 1))
            ],
            [
                PySimpleGUI.Text('Название листа в книге Excel: ', size=(23, 1)),
                PySimpleGUI.InputText(key='sheet', size=(62, 1))
            ],
            [
                PySimpleGUI.Text('Путь к папке со сканами: ', size=(19, 1)),
                PySimpleGUI.InputText(key='folder', size=(56, 1)),
                PySimpleGUI.FolderBrowse(target='folder', size=(7, 1))
            ],
            [
                PySimpleGUI.Output(size=(88, 20))
            ],
            [
                PySimpleGUI.Submit(button_text='Start'),
                PySimpleGUI.Cancel(button_text='Exit'),
            ]
        ]

        return PySimpleGUI.Window('Hyperlinks creator', layout)

    def font_size_menu(self):
        settings = self.load_settings()
        font_list = settings['font_list'].split(',')
        font_name = settings['font_name']
        font_size = int(settings['font_size'])

        layout = [
            [PySimpleGUI.Combo(font_list, default_value=font_name)],
            [PySimpleGUI.Spin([sz for sz in range(10, 21)], font=('Helvetica 20'), initial_value=font_size,
                              change_submits=True,
                              key='spin'),
             PySimpleGUI.Slider(range=(10, 20), orientation='h', size=(10, 25),
                                change_submits=True, key='slider', font=('Helvetica 20'), default_value=font_size),
             PySimpleGUI.Text("Ab", size=(2, 1), font="Helvetica " + str(font_size), key='text')],
            [
                PySimpleGUI.Submit(button_text='Ok'),
                PySimpleGUI.Cancel(button_text='Cancel')
            ]
        ]

        sz = font_size
        window = PySimpleGUI.Window("Font size selector", layout, grab_anywhere=False)
        # Event Loop
        while True:
            event, values = window.read()
            if event in (PySimpleGUI.WIN_CLOSED, 'Cancel'):
                window.close()
                break
            sz_spin = int(values['spin'])
            sz_slider = int(values['slider'])
            sz = sz_spin if sz_spin != font_size else sz_slider
            if sz != font_size:
                font_size = sz
                font = "Helvetica " + str(font_size)
                window['text'].update(font=font)
                window['slider'].update(sz)
                window['spin'].update(sz)
            if event in 'Ok':
                self.save_settings({'font_size': font_size})
                window.close()
                break


if __name__ == '__main__':
    gg = GUI()
    window_main = gg.main_menu()

    while True:                             # The Event Loop
        event, values = window_main.read()
        # print(event, values) #debug

        if event in (None, 'Exit', 'Cancel'):
            break

        elif event in 'Font':
            print('Font')

        elif event in 'Font size':
            # window_main.hide()
            gg.font_size_menu()
            # window_main.UnHide()

        elif event == 'Hyperlink color':
            pass

        elif event == 'Start':
            print('Start')
