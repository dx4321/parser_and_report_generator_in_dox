import locale
import math
import datetime

import calendar
from docxtpl import DocxTemplate
import socks
import socket
import gspread
from googleapiclient.discovery import build
from gspread import Worksheet
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials

from utils.utils import Config, get_config


class Google:
    def __init__(self, con: Config):
        self.config = con

        # настройки прокси-сервера
        if self.config.proxy:
            # установка соединения с прокси сервером
            socks.setdefaultproxy(
                socks.PROXY_TYPE_HTTP,
                self.config.proxy.proxy_host,
                self.config.proxy.proxy_port,
                True,
                self.config.proxy.proxy_user,
                self.config.proxy.proxy_pass
            )
            self.socket = socket
            self.socket.socket = socks.socksocket

        # учетные данные для доступа к Google API
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            con.oauth2client_service_account_file,
            ['https://www.googleapis.com/auth/spreadsheets',
             'https://www.googleapis.com/auth/drive'])

        # подключение к таблице Google Sheets
        self.client = gspread.authorize(creds)

        self.url = con.google_url
        self.con_url = self.google_con_url

    def close(self):
        socks.setdefaultproxy()
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    @staticmethod
    def get_current_data():
        """ Получить текущую дату в формате 18.04.2023 """

        return datetime.date.today()

    def convert_datetime_for_google_format(self):
        """ Конвертировать текущую дату в гугл формат """
        return self.get_current_data().strftime('%Y.%m.%d')

    @property
    def google_con_url(self):
        return self.client.open_by_url(self.url)


config = get_config('config.yaml')
google = Google(config)
cur_data = google.convert_datetime_for_google_format()
month_year = "05.23"
month = month_year.split(".")[0]
worksheet: Worksheet = google.con_url.worksheet(month_year)
print(worksheet.title)

data = worksheet.get_all_values()

work_data = []
shapka = []

# Находим индекс значения со словом "Итого"
index_itogo = next((i for i, item in enumerate(data[0]) if item.startswith("Итого")), -1)
index_a = next((i for i, item in enumerate(data[0]) if item == "А"), -1)
index_o = next((i for i, item in enumerate(data[0]) if item == "О"), -1)
index_zashet_o = next((i for i, item in enumerate(data[0]) if item == "За счёт О"), -1)
index_first_ne_pust = next((i for i, item in enumerate([row[0] for row in data]) if item == "" and i != 0), -1)

# Определение диапазона для получения ячеек
start_row = 1
end_row = 21

# Получение форматирования ячеек с использованием Google Sheets API
credentials = Credentials.from_service_account_file(config.oauth2client_service_account_file)

service = build('sheets', 'v4', credentials=credentials)
sheet_id = google.con_url.id
range_string = f"{worksheet.title}!A{start_row}:AK{end_row}"
response = service.spreadsheets().get(spreadsheetId=sheet_id, ranges=[range_string], includeGridData=True).execute()
cell_data = response['sheets'][0]['data'][0]['rowData']

users = [row[0] for row in data][1:index_first_ne_pust]  # Получаем список сотрудников


class Colors:
    """ Цвета в hex """
    silver = "#C0C0C0"
    white = "#FFFFFF"


class RGBColors:
    """ Цвета в RGB """
    blue = (0.2901961, 0.5254902, 0.9098039)
    red = (0.8352941, 0.6509804, 0.7411765)
    silver = (0.8, 0.8, 0.8)


for i, (row, row_full) in enumerate(zip(data, cell_data)):
    if i == 0:
        shapka = row[:index_itogo][1:]
        continue
    # Нужно вывести до последней не пустой ячейки
    if row[0] == "":
        break

    for j, (val, cell) in enumerate(zip(row, row_full['values'])):
        if j == 0:
            # Если первый столбец, то присвоить значение
            row[j] = val
            continue

        cell_format = cell.get('effectiveFormat', {})
        cell_background_color = cell_format.get('backgroundColor', {})
        cell_color = cell_background_color.get('red', 0), cell_background_color.get('green',
                                                                                    0), cell_background_color.get(
            'blue', 0)

        rgb = cell_color

        if cell_color == RGBColors.blue:
            # Если синий то + *
            row[j] = {
                "val": f"{val}*",
                "bg": Colors.white
            }
            continue
        if cell_color == RGBColors.red:
            # Если красный то + **
            row[j] = {
                "val": f"{val}**",
                "bg": Colors.white
            }
            continue
        if cell_color == RGBColors.silver:
            row[j] = {
                "val": f"{val}",
                "bg": Colors.silver
            }
            continue

        if j == index_o:
            # Сложить 2 столбца "О и За счёт О" и сделать 1 столбец
            # Удаляем 1 столбец За счет О, вместо него суммируем за "счет О" с "О"
            o1 = val
            o2 = row[index_zashet_o]

            # Если в строке есть символ разделитель, делаем так, что бы
            # мы могли конвертировать число в float для суммирования
            if o1 == "":
                o1 = 0
            elif ',' in o1 or '.' in o1:
                o1 = float(o1.replace(',', '.'))
            else:
                o1 = int(o1)

            if o2 == "":
                o2 = 0
            elif ',' in o2 or '.' in o2:
                o2 = float(o2.replace(',', '.'))
            else:
                o2 = int(o2)

            solo = o1 + o2

            if solo == 0:
                solo = ""
            row[index_zashet_o] = solo

            row.pop(j)  # Используем j вместо i
            continue
        if j >= index_a:
            # Если столбец с индексом, а и больше то
            row[j] = val
            continue
        row[j] = {
            "val": val,
            "bg": Colors.white
        }
    work_data.append(row[:index_itogo])


work_data = [row[1: index_itogo] for row in work_data]
# Обновить индексы
index_itogo = index_itogo - 2
index_a = index_a - 1
index_o = index_o - 1

tpl = DocxTemplate('templates/основной.docx')

# Устанавливаем русскую локаль
locale.setlocale(locale.LC_ALL, ('ru_RU', 'UTF-8'))

# Номер месяца
month_number = int(month)

# Получаем русское название месяца с большой буквы
month_name = calendar.month_name[month_number].capitalize()

context = {'month': month_name, 'year': "3", 'col_labels': shapka[:index_a]}

work_data_bez_abo = [row[: index_a] for row in work_data]

# Разделение tbl_contents на две части
tbl_contents = [
    {
        "label": str(val),
        "cols": [data for data in work_data_bez_abo[i]],
        "a": work_data[i][index_a],
        "b": work_data[i][index_a + 1],
        "o": work_data[i][index_a + 2],  # Сложить две ячейки индекс а + 2 и индекс а + 3
        "it": work_data[i][index_itogo]  #
    } for i, val in enumerate(users)
]

users_count_for_2 = math.ceil(len(users) / 2)

tbl_contents_part1 = tbl_contents[:users_count_for_2]
tbl_contents_part2 = tbl_contents[users_count_for_2:]


# Добавление tbl_contents_part1 и tbl_contents_part2 в context
context['tbl_contents_part1'] = tbl_contents_part1
context['tbl_contents_part2'] = tbl_contents_part2

tpl.render(context)

tpl.save(f'output/отчет за {month_year.replace(".", "_")}.docx')
