from openpyxl import load_workbook
from get_messages import data
from time import ctime


def get_date(data):
    values = []
    for string in data:
        if string.startswith('@'):
            value = string.split('@')
            values.append([value[1], value[2]])
    return values


def get_last_value():
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
            with open('log.txt', 'a') as file:
                file.write(str(send_value[0]))
                file.write(' ')
                file.write(ctime())
                file.write('\n')
        else:
            with open('log.txt', 'a') as file:
                file.write(f'Data is already in file! - {send_value[0]}')
    book.save(name_book)
    book.close()


def delete_data():
    today_date = ctime().split()
    day = today_date[0]
    time = today_date[4].split(':')
    if day == 'Sun' and int(time[0]) > 20:
        name_book = 'coasts.xlsx'
        book = load_workbook(name_book)
        sheet = book['data']
        sheet.delete_rows(3, 50)
        book.save(name_book)
        book.close()

values = get_date(data)
last_values = get_last_value()
