from openpyxl import load_workbook
from get_messages import data


def get_date(data):
    values = []
    for string in data:
        if string.startswith('@'):
            value = string.split('@')
            values.append([value[1], value[2]])
    return values


def open_book():
    last_value = []
    name_book = 'coasts.xlsx'
    book = load_workbook(name_book)
    worksheet = book.active
    for col in worksheet.iter_cols(1, 2):
        last_value.append(col[worksheet.max_row - 1].value)
    book.close()
    return last_value


def write_data(values, last_value):
    name_book = 'coasts.xlsx'
    book = load_workbook(name_book)
    sheet = book['data']
    for value in values:
        send_value = [int(value[0]), value[1]]
        if last_value != send_value:
            sheet.append(send_value)
        else:
            print('Error')
    book.save(name_book)
    book.close()


values = get_date(data)
last_values = open_book()
write_data(values, last_values)
