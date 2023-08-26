from openpyxl import load_workbook
from get_messages import data


def get_date(values):
    numbers = []
    for string in values:
        if string.startswith('@'):
            value = string.split('@')
            numbers.append(value[1])
    print(numbers)


# name_book = 'coasts.xlsx'
# book = load_workbook(name_book)
# sheet = book['data']
# sheet.append([6000])
# book.save(name_book)
# book.close()

get_date(data)
