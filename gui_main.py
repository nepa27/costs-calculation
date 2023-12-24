from tkinter import LEFT, Menu, RIGHT, Tk, Y
from tkinter.ttk import Treeview, Scrollbar
from change_excel import (delete_data,
                          delete_data_in_the_weekend,
                          get_data_from_email,
                          get_last_costs,
                          summ_value_a_week,
                          write_data)
from get_messages import get_messages_from_email


class Window:
    """ Create window and work with data. """

    NAME_COLUMNS: tuple = ('Сколько потрачено',
                           'На что потрачено',
                           'Сумма трат за неделю')

    def __init__(self):
        self.scrollbar = None
        self.menu = None
        self.table = None
        self.window = None

    def create_window(self, width_window, height_window):
        """ Create main window. """
        self.window = Tk()
        self.window.title('My costs')
        # self.window.iconbitmap(default='1.ico')
        self.window['bg'] = '#999'
        self.window.geometry(str(width_window) + 'x' + str(height_window))

        self.menu = Menu(self.window)
        self.window.configure(menu=self.menu)
        self.menu.add_command(label='Добавить',
                              command=lambda: self.append_item())
        self.menu.add_command(label='Обновить',
                              command=lambda: self.create_table())
        self.menu.add_command(label='Удалить',
                              command=lambda: self.delete_item())
        self.menu.add_command(label='О программе',
                              command=lambda: print('Info'))
        self.menu.add_command(label='Выход',
                              command=lambda: self.destroy())
        self.create_table()
        self.window.mainloop()

    def create_table(self):
        """ Create table. """
        if self.table:
            self.table.destroy()
        self.table = Treeview(self.window,
                              columns=('Сколько потрачено',
                                       'На что потрачено',
                                       'Сумма трат за неделю'),
                              show='headings',
                              )
        [self.table.heading(column_name, text=column_name)
         for column_name in self.NAME_COLUMNS]
        [self.table.insert(parent='',
                           index=0,
                           values=(costs, name_costs))
         for costs, name_costs in get_last_costs()]
        self.table.insert(parent='',
                          index=0,
                          values=('', '', summ_value_a_week(get_last_costs())))
        self.table.bind('<<TreeviewSelect>>', self.item_select)
        # self.table.bind('<Delete>', self.delete_item)

        self.window.update()
        self.table.pack(side=LEFT, fill='both', expand=True)
        if self.scrollbar:
            self.scrollbar.destroy()
        self.scrollbar = Scrollbar(orient="vertical", command=self.table.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.table['yscrollcommand'] = self.scrollbar.set

    def item_select(self, _):
        """ Print name item which focus. """
        print(self.table.selection())
        for item in self.table.selection():
            print(self.table.item(item)['values'])

    def append_item(self):
        """ Delete item into the table. """
        print('Add', self.menu)

    def delete_item(self):
        """ Delete item which focus. """
        delete_data(self.table.selection())
        for item in self.table.selection():
            self.table.delete(item)

    def destroy(self):
        """ Destroy window. """
        self.window.destroy()


if __name__ == '__main__':
    new_values = get_data_from_email(get_messages_from_email())
    last_values = get_last_costs()

    write_data(new_values, last_values)
    delete_data_in_the_weekend(last_values)

    window = Window()
    window.create_window(800, 300)
