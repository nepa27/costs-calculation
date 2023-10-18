from openpyxl import load_workbook
from get_messages import data
from datetime import datetime as dt
from time import ctime

NAME_BOOK = 'costs.xlsx'


def get_date(data: list) -> list:
    values = []
    for string in data:
        if string.startswith('@'):
            value = string.split('@')
            values.append((int(value[1]), value[2]))
    return values


def get_last_value() -> list:
    name = []
    value = []
    book = load_workbook(NAME_BOOK)
    sheet = book.active
    for col in range(1, 3):
        for row in range(3, sheet.max_row + 1):
            if col == 1:
                name.append(sheet.cell(row, col).value)
            if col == 2:
                value.append(sheet.cell(row, col).value)
    last_value = list(zip(name, value))
    return last_value


def write_data(values: list, last_value: list):
    book = load_workbook(NAME_BOOK)
    sheet = book['data']
    for value in values:
        number_value = str(value[0])
        name_value = value[1]
        time = dt.today().strftime("Date: %d.%m.%Y | Time: %H:%M:%S |")
        if value in last_value:
            pass
        else:
            sheet.append(value)
            with open('log.txt', 'a') as file:
                file.write(f'| {number_value} р.'.ljust(11))
                file.write(f'expended for "{name_value}" |')
                file.write(' ')
                file.write(str(time).rjust(22))
                file.write('\n')

    book.save(NAME_BOOK)
    book.close()


def summ_value_a_week(values_and_numbers: list) -> int:
    summ_coasts = 0
    for value_and_number in values_and_numbers:
        summ_coasts += value_and_number[0]
    return summ_coasts


def delete_data(last_value: list):
    today_date = ctime().split()
    day = today_date[0]
    time = today_date[3].split(':')
    if day == 'Sun' and int(time[0]) > 22:
        result_summ = summ_value_a_week(last_value)
        book = load_workbook(NAME_BOOK)
        sheet = book['data']
        sheet.delete_rows(3, 50)
        book.save(NAME_BOOK)
        book.close()
        with open('log.txt', 'a') as file:
            file.write(str(f'Summ for a week: {result_summ}'))
            file.write(' ')
            file.write(ctime())
            file.write('\n')


values = get_date(data)
last_values = get_last_value()