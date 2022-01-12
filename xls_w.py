import xlwings
import time


class Excel:
    def __init__(self, registry_path, dir_scan, ws_name, settings):
        self.font_size = round(int(settings['font_size']), 1)
        self.font_name = settings['font_name']
        self.hyperlink_color = settings['hyperlink_color']
        self.dir_scan = dir_scan

        if registry_path not in [True, None, 'None', '']:
            self.wb = xlwings.Book(registry_path)
        else:
            self.wb = xlwings.books.active

        if ws_name not in [True, None, 'None', '']:
            self.ws_name = ws_name
        else:
            self.ws_name = self.wb.sheets.active.name

        print(f'Документ {self.wb.name} открыт.')
        if self.ws_name not in list(sheet.name for sheet in self.wb.sheets):
            print(f'Лист {self.ws_name} отсутствует в книге.')
        self.ws = self.wb.sheets[self.ws_name]
        print(f'Выбран лист {self.ws_name}.')

    def get_column(self):
        list_column = list(str(self.ws[f'A{i}'].value)[:str(self.ws[f'A{i}'].value).rfind('.') if '.' in str(self.ws[f"A{i}"].value) else None] for i in range(3, self.ws.range('A1').end('down').row + 1))
        list_column = list(str(i) for i in list_column if i is not None)
        print('Получен список регистрационных номеров из столбца "А".')
        return list_column

    def get_path_active_book(self):
        return self.wb.fullname

    def create_hyperlinks(self, name, link_name, position):
        self.ws[f'H{position}'].add_hyperlink(f'{self.dir_scan}\\{link_name}', name)
        self.ws[f'H{position}'].font.name = self.font_name
        self.ws[f'H{position}'].font.size = self.font_size
        # if self.hyperlink_color != 'None':
        self.ws[f'H{position}'].font.color = self.hyperlink_color
        self.ws[f'H{position}'].api.HorizontalAlignment = -4108
        self.borders_all(f'H{position}')

    def borders_all(self, cell):
        for i in range(7, 13):
            self.ws[cell].api.Borders(i).LineStyle = 1

    # def save(self):
    #     self.wb.save(self.registry_path)

    @staticmethod
    def hex_to_rgb(hex_color):
        # if hex_color == 'None':
        #     return tuple([5, 99, 193])
        h = hex_color.lstrip('#')
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

    def check_hyperlink(self, name, link_name, position):
        try:
            hyperlink = self.ws[f'H{position}'].hyperlink
        except:
            hyperlink = None

        # print(name, self.ws[f'H{position}'].value)
        # print(f'{self.dir_scan}\\{link_name}'.replace('\\', '/'), hyperlink.replace('\\', '/'))
        # print(self.font_name, self.ws[f'H{position}'].font.name)
        # print(self.font_size, self.ws[f'H{position}'].font.size)
        # print(Excel.hex_to_rgb(self.hyperlink_color), self.ws[f'H{position}'].font.color)

        if (
                name == self.ws[f'H{position}'].value
                and f'{self.dir_scan}\\{link_name}'.replace('\\', '/') == hyperlink.replace('\\', '/')
                and self.font_name == self.ws[f'H{position}'].font.name
                and self.font_size == self.ws[f'H{position}'].font.size
                and Excel.hex_to_rgb(self.hyperlink_color) == self.ws[f'H{position}'].font.color
        ):
            return True
        else:
            return False

        # time.sleep(10)


if __name__ == '__main__':
    pass
    import os
    from GUI import GUI


    registry_path = r'C:/Users/suhorukov.iv/Desktop/Реестр входящих 2020-2021.xlsx'
    dir_scan = r'//fs.corp.skaarena.ru/SHARE/Documents/OTDEL-SECRETARY/Регистрация документов/ИСХОДЯЩИЕ 2021'
    ws_name = '2021'

    # position = os.path.abspath(__file__).rfind('\\')
    # config_path = f'{os.path.abspath(__file__)[:position]}\\config.ini'
    gg = GUI()
    settings = gg.load_settings()
    xxl = Excel(registry_path, dir_scan, ws_name, settings)
    print(xxl.get_column())
