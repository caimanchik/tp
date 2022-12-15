import csv
import math
import re
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import os
import pdfkit

from jinja2 import Environment, FileSystemLoader
from functools import reduce
from datetime import datetime
from typing import List, Dict, Tuple
from openpyxl import Workbook
from openpyxl.styles import Side, Border, Font
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from openpyxl.utils import get_column_letter


class Salary(object):
    __currency_to_rub = {
        "AZN": 35.68,
        "BYR": 23.91,
        "EUR": 59.90,
        "GEL": 21.74,
        "KGS": 0.76,
        "KZT": 0.13,
        "RUR": 1,
        "UAH": 1.64,
        "USD": 60.66,
        "UZS": 0.0055,
    }

    def __init__(self, values: List[str]):
        [self.__salary_from, self.__salary_to, self.__salary_currency] = values

    def __float__(self):
        return (float(self.__salary_from) + float(self.__salary_to)) / 2 * self.__currency_to_rub[
            self.__salary_currency.upper()]


class Vacancy:

    def __init__(self, row: List[str], title: List[str]):
        self.__name = None
        self.__salary = None
        self.__area_name = None
        self.__published_at = None
        self.__salary_from = None
        self.__salary_to = None
        self.__salary_currency = None

        fields_cases = {
            'name': lambda value: self.__set_value('name', HelpMethods.delete_rubbish(value)),
            'salary_from': lambda value: self.__set_value('salary_from', HelpMethods.delete_rubbish(value)),
            'salary_to': lambda value: self.__set_value('salary_to', HelpMethods.delete_rubbish(value)),
            'salary_currency': lambda value: self.__set_value('salary_currency', HelpMethods.delete_rubbish(value)),
            'area_name': lambda value: self.__set_value('area_name', HelpMethods.delete_rubbish(value)),
            'published_at': lambda value: self.__set_value(
                'published_at',
                Vacancy.__get_date(HelpMethods.delete_rubbish(value))),
        }

        for i, field in enumerate(row):
            if title[i] not in fields_cases:
                continue

            fields_cases[title[i]](field)

        self.__set_salary()

    def get_salary(self) -> float:
        return float(self.__salary)

    def get_date(self) -> str:
        return self.__published_at

    def get_area(self) -> str:
        return self.__area_name

    def is_suitible(self, name: str) -> bool:
        return self.__name.count(name) > 0

    def __set_value(self, key, value):
        self.__dict__['_Vacancy__' + key] = value

    def __set_salary(self):
        self.__salary = Salary([self.__salary_from, self.__salary_to, self.__salary_currency])

    @staticmethod
    def __get_date(date: str) -> str:
        return str(datetime.fromisoformat(date[:-2] + ":" + date[-2:]).year)


class DataSet:

    def __init__(self, file_name):
        self.__file_name = file_name
        self.__vacancies_objects: List[Vacancy] = []
        self.__title = None
        self.__vacancies_years = {}
        self.__vacancies_areas = {}
        self.__len = 0

        with open(file_name, mode='r', encoding='utf-8-sig') as vacancies:
            file_reader = csv.reader(vacancies, delimiter=",")
            has_title = False

            for row in file_reader:
                if not has_title:
                    self.__title = row
                    has_title = True
                    continue

                if row.count('') != 0 or len(row) < len(self.__title) - 1: continue

                self.__validate_vacancy(row)
                self.__len += 1

    def __check_dataset(self):
        if self.__title is None:
            HelpMethods.quit_program('Пустой файл')
        if len(self.__vacancies_years) == 0:
            HelpMethods.quit_program('Нет данных')

    def __validate_vacancy(self, row: List[str]):
        vacancy = Vacancy(row, self.__title)

        now_date = self.__vacancies_years.get(vacancy.get_date(), [])
        now_date.append(vacancy)
        self.__vacancies_years[vacancy.get_date()] = now_date

        now_area = self.__vacancies_areas.get(vacancy.get_area(), [])
        now_area.append(vacancy)
        self.__vacancies_areas[vacancy.get_area()] = now_area

    def get_vacancies_years(self, func=None) -> Dict[str, List[int]]:
        if func is None:
            return DataSet.get_structured_salaries(self.__vacancies_years)

        result = {}

        for year in self.__vacancies_years.keys():
            result[year] = []
            for vacancy in self.__vacancies_years[year]:
                if func(vacancy):
                    result[year].append(vacancy)

        return DataSet.get_structured_salaries(result)

    def get_vacancies_cities(self) -> Tuple[List[List[float]], List[List[int]]]:
        return DataSet.get_structured_cities(self.__vacancies_areas, self.__len)

    @staticmethod
    def get_structured_salaries(vacancies: Dict[str, list]) -> Dict[str, List[int]]:
        salaries = {}

        for i, year in enumerate(vacancies.keys()):

            summ = 0
            for vacancy in vacancies[year]:
                summ += vacancy.get_salary()

            salaries[year] = [
                math.floor(summ / len(vacancies[year])) if len(vacancies[year]) > 0 else 0,
                len(vacancies[year])
            ]

        return salaries

    @staticmethod
    def get_structured_cities(vacancies: Dict[str, list], l: int) -> Tuple[List[List[float]], List[List[int]]]:

        cities_s = []
        fract = []

        for key, value in vacancies.items():

            percent = round(len(value) / l, 4)
            if percent < 0.01:
                continue

            summ = 0
            for vacancy in vacancies[key]:
                summ += vacancy.get_salary()

            cities_s.append([key, math.floor(summ / len(value))])
            fract.append([key, percent])

        fract.sort(key=lambda x: x[1], reverse=True)
        cities_s.sort(key=lambda x: x[1], reverse=True)

        return fract, cities_s


class InputConnect:

    def __init__(self):
        self.file_name = None
        self.vacancy = None

    def read_console(self):
        # if input() == '':
        #     self.file_name = '../vacancies.csv'
        #     self.vacancy = 'Аналитик'
        # else:
        #     self.file_name = input("Введите название файла: ")
        #     self.vacancy = input("Введите название профессии: ")
        self.file_name = input("Введите название файла: ")
        self.vacancy = input("Введите название профессии: ")

    @staticmethod
    def write_console(s_all, s_filtered, fract, cities_s):

        InputConnect.__write_salaries(s_all)
        InputConnect.__write_salaries(s_filtered, ' для выбранной профессии')

        InputConnect.__write_salaries_cities(fract, cities_s)

    @staticmethod
    def __write_salaries(salaries: Dict[str, List[int]], sufix=''):
        s = f'Динамика уровня зарплат по годам{sufix}' + ': {'

        print(s, end='')

        for i, year in enumerate(salaries.keys()):
            if i != 0:
                print(', ', end='')

            print(f'{year}: {salaries[year][0]}', end='')

        print('}')

        s = f'Динамика количества вакансий по годам{sufix}' + ': {'

        print(s, end='')

        for i, year in enumerate(salaries.keys()):
            if i != 0:
                print(', ', end='')

            print(f'{year}: {salaries[year][1]}', end='')

        print('}')

    @staticmethod
    def __write_salaries_cities(fract: List[List[float]], cities_s: List[List[int]]):

        print('Уровень зарплат по городам (в порядке убывания): {', end='')
        for i, e in enumerate(cities_s[:10]):
            if i != 0:
                print(', ', end='')
            print(f"'{e[0]}': {e[1]}", end='')

        print('}')

        print('Доля вакансий по городам (в порядке убывания): {', end='')
        for i, e in enumerate(fract[:10]):
            if i != 0:
                print(', ', end='')
            print(f"'{e[0]}': {e[1]}", end='')

        print('}')


class HelpMethods:

    @staticmethod
    def quit_program(message):
        print(message)
        quit()

    @staticmethod
    def delete_rubbish(s: str) -> str:
        rubbish_html = re.compile('<.*?>')

        return ' '.join(re.sub(rubbish_html, '', s).split()).strip()


class report:

    def __init__(self, vacancy: str,
                 s_all: Dict[str, List[int]],
                 s_filtered: Dict[str, List[int]],
                 fract: List[List[float]],
                 cities_s: List[List[int]]):
        self.__salaries_all = s_all
        self.__salaries_filtered = s_filtered
        self.__fraction = fract
        self.__cities_salaries = cities_s
        self.__vacancy = vacancy

        self.__names_ws1 = {
            'A1': 'Год',
            'B1': 'Средняя зарплата',
            'C1': f'Средняя зарплата - {vacancy}',
            'D1': 'Количество вакансий',
            'E1': f'Количество вакансий - {vacancy}',
        }

        self.__names_ws2 = {
            'A1': 'Город',
            'B1': 'Уровень зарплат',
            'D1': 'Город',
            'E1': 'Доля вакансий'
        }
        self.__titles = [
            'Уровень зарплат по годам',
            'Количество вакансий по годам',
            'Уровень зарплат по городам',
            'Доля вакансий по городам',
        ]

    def generate_excel(self):
        wb = Workbook()

        ws1 = wb.active
        ws2 = wb.create_sheet('Статистика по городам')

        report.__make_ws1(ws1, self.__salaries_all, self.__salaries_filtered, self.__names_ws1)
        report.__make_ws2(ws2, self.__fraction, self.__cities_salaries, self.__names_ws2)

        wb.save('report.xlsx')

    def generate_image(self):
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(nrows=2, ncols=2)

        self.__create_bar(
            ax1,
            self.__salaries_all,
            self.__salaries_filtered,
            0,
            ['Средняя з/п', f'З/п {self.__vacancy}'],
            self.__titles[0]
        )

        self.__create_bar(
            ax2,
            self.__salaries_all,
            self.__salaries_filtered,
            1,
            ['Количество вакансий', f'Количество вакансий {self.__vacancy}'],
            self.__titles[1]
        )

        self.__create_barh(ax3, self.__cities_salaries[:10], self.__titles[2])
        self.__create_pie(ax4, self.__fraction[:10], self.__titles[3])

        fig.tight_layout()
        fig.set_size_inches(8, 6)
        fig.set_dpi(300)
        fig.savefig('graph.png', dpi=300)

        plt.show()

    def generate_pdf(self):
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template("template.html")
        config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
        path = os.path.abspath('')
        rows_1 = self.__generate_rows_1(self.__salaries_all, self.__salaries_filtered)
        rows_2, rows_3 = self.__generate_rows_23(self.__fraction, self.__cities_salaries)

        context = {
            'path': path,
            'vacancy': self.__vacancy,
            'rows1': rows_1,
            'rows2': rows_2,
            'rows3': rows_3,
        }

        pdf_template = template.render(context)
        pdfkit.from_string(pdf_template, 'report.pdf', configuration=config, options={"enable-local-file-access": None})

    @staticmethod
    def __generate_rows_1(s_all: Dict[str, List[int]], s_filtered: Dict[str, List[int]]) -> List[Dict[str, str | int]]:
        rows = []
        for key in s_all.keys():
            row = {
                'year': key,
                'average': s_all[key][0],
                'average_v': s_filtered[key][0],
                'count': s_all[key][1],
                'count_v': s_filtered[key][1],
            }

            rows.append(row)

        return rows

    @staticmethod
    def __generate_rows_23(fract: List[List[float]], cities_s: List[List[int]]) \
            -> Tuple[List[Dict[str, int]], List[Dict[str, float]]]:
        count = 10
        rows_2 = []
        rows_3 = []

        for i in range(count):
            row = {
                'city': cities_s[i][0],
                'salary': cities_s[i][1]
            }

            rows_2.append(row)

            row = {
                'city': fract[i][0],
                'fraction': str(round(fract[i][1] * 100, 2)) + '%'
            }
            rows_3.append(row)

        return rows_2, rows_3

    @staticmethod
    def __create_bar(
            ax,
            data1: Dict[str, List[int]],
            data2: Dict[str, List[int]],
            index: int,
            legend: List[str],
            title: str
    ):
        width = 0.35
        labels_x = data1.keys()
        first = report.__get_data(data1, index)
        second = report.__get_data(data2, index)
        points = range(len(labels_x))

        ax.bar(list(map(lambda x: x - width / 2, points)), first, width, label=legend[0])
        ax.bar(list(map(lambda x: x + width / 2, points)), second, width, label=legend[1])
        ax.set_title(title)
        ax.legend(prop={'size': 8})
        ax.grid(axis='y')

        for label in ax.get_yticklabels():
            label.set_fontsize(8)

        ax.set_xticks(points, labels_x, fontsize=8, rotation=90)

    @staticmethod
    def __create_barh(ax, data: List[List[float]], title: str):
        cities = list(map(lambda x: report.__refactor_label(x[0]), data))
        y_pos = list(range(len(cities)))

        ax.barh(y_pos, list(map(lambda x: x[1], data)), align='center')

        ax.set_yticks(y_pos, labels=cities, fontsize=6)
        ax.invert_yaxis()
        ax.grid(axis='x')

        for label in ax.get_xticklabels():
            label.set_fontsize(8)

        ax.set_title(title)

    @staticmethod
    def __create_pie(ax, data: List[List[float]], title: str):
        cities = list(map(lambda x: x[0], data)) + ['Другие']
        others = 1 - reduce(lambda x, y: x + y[1], data, 0)

        ax.pie(list(map(lambda x: x[1], data)) + [others],
               labels=cities, textprops={'size': 6}, colors=mcolors.BASE_COLORS)

        ax.set_title(title)

    @staticmethod
    def __refactor_label(label: str) -> str:
        spaces = re.compile('\s+')
        line = re.compile('-+')

        label = re.sub(spaces, '\n', label)
        return re.sub(line, '-\n', label)

    @staticmethod
    def __get_data(data: Dict[str, List[int]], i: int) -> List[int]:
        return list(map(lambda x: x[i], data.values()))

    @staticmethod
    def __make_ws1(
            ws,
            s_all: Dict[str, List[int]],
            s_filtered: Dict[str, List[int]],
            title: Dict[str, str]
    ):
        ws.title = 'Статистика по годам'
        report.__create_title(ws, title)

        for key in s_all.keys():
            ws.append([
                int(key),
                s_all[key][0],
                s_filtered[key][0],
                s_all[key][1],
                s_filtered[key][1],
            ])

        report.__set_border(ws, f'A1:E{len(s_all) + 1}')
        report.__refactor_rows(ws)

    @staticmethod
    def __make_ws2(
            ws,
            fract: List[List[float]],
            cities_s: List[List[int]],
            title: Dict[str, str]
    ):
        report.__create_title(ws, title)
        count = 10

        for i in range(count):
            row = []

            row += [cities_s[i][0], cities_s[i][1]] if len(cities_s) >= i + 1 else ['', '']
            row += ['']
            row += [fract[i][0], fract[i][1]] if len(fract) >= i + 1 else ['', '']

            ws.append(row)

        report.__add_percentage(ws, count, 'E')
        report.__set_border(ws, f'A1:B{count + 1}')
        report.__set_border(ws, f'D1:E{count + 1}')
        report.__refactor_rows(ws)

    @staticmethod
    def __add_percentage(ws, count: int, column: str):
        for i in range(2, count + 2):
            ws[f'{column}{i}'].number_format = FORMAT_PERCENTAGE_00

    @staticmethod
    def __set_border(ws, cell_range):
        line = Side(border_style="thin", color="000000")
        border = Border(top=line, left=line, right=line, bottom=line)

        for row in ws[cell_range]:
            for cell in row:
                cell.border = border

    @staticmethod
    def __create_title(ws, title: Dict[str, str]):
        font = Font(bold=True)

        for key, value in title.items():
            ws[key] = value
            ws[key].font = font

    @staticmethod
    def __refactor_rows(ws):
        for i, col in enumerate(ws.iter_cols()):
            l = 0
            for cell in col:
                v = cell.value if cell.value is not None else ''
                l = max(l, len(str(v)))

            ws.column_dimensions[get_column_letter(i + 1)].width = l + 3 if l != 0 else 0


connect = InputConnect()
connect.read_console()

dataset = DataSet(connect.file_name)

salaries_all = dataset.get_vacancies_years()
salaries_filtered = dataset.get_vacancies_years(lambda x: x.is_suitible(connect.vacancy))
fraction, cities_salaries = dataset.get_vacancies_cities()

connect.write_console(salaries_all, salaries_filtered, fraction, cities_salaries)

rep = report(connect.vacancy,
             salaries_all,
             salaries_filtered,
             fraction,
             cities_salaries
             )

rep.generate_excel()
rep.generate_image()
rep.generate_pdf()