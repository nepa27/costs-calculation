from openpyxl import load_workbook
from get_messages import data
from time import ctime


def get_date(data):
    values = []
    for string in data:
        if string.startswith('@'):
            value = string.split('@')
            values.append((int(value[1]), value[2]))
    return values


def get_last_value():
    name = []
    value = []
    name_book = 'coasts.xlsx'
    book = load_workbook(name_book)
    sheet = book.active
    for col in range(1, 3):
        for row in range(3, sheet.max_row+1):
            if col == 1:
                name.append(sheet.cell(row, col).value)
            if col == 2:
                value.append(sheet.cell(row, col).value)
    last_value = list(zip(name, value))
    return last_value


def write_data(values, last_value):
    name_book = 'coasts.xlsx'
    book = load_workbook(name_book)
    sheet = book['data']
    for value in values:
        if value in last_value:
            pass
        else:
            sheet.append(value)
            with open('log.txt', 'a') as file:
                file.write(str(value[0]))
                file.write(' ')
                file.write(ctime())
                file.write('\n')

    book.save(name_book)
    book.close()


def delete_data():
    today_date = ctime().split()
    day = today_date[0]
    time = today_date[4].split(':')
    if day == 'Sun' and int(time[0]) > 22:
        name_book = 'coasts.xlsx'
        book = load_workbook(name_book, data_only=True)
        worksheet = book.active
        summ = worksheet['C2'].value
        with open('log.txt', 'a') as file:
            file.write(str(f'Summ for a week: {str(summ)}'))
            file.write(' ')
            file.write(ctime())
            file.write('\n')
        sheet = book['data']
        sheet.delete_rows(3, 50)
        book.save(name_book)
        book.close()


values = get_date(data)
last_values = get_last_value()
