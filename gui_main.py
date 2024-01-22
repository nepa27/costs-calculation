import tkinter
import tkinter.ttk
from tkinter import messagebox
from re import match

import change_excel
import get_messages


class Window:
    """ Create window and work with data. """

    NAME_COLUMNS: tuple = ('Сколько потрачено',
                           'На что потрачено',
                           'Сумма трат за неделю')

    def __init__(self):
        self.check = None
        self.name_label = None
        self.value_label = None
        self.frame_bottom = None
        self.frame_top = None
        self.top = None
        self.left = None
        self.display_height = None
        self.display_width = None
        self.info_message = None
        self.add_button = None
        self.cost_value = None
        self.cost_name = None
        self.append_window = None
        self.scrollbar = None
        self.menu = None
        self.table = None
        self.window = None

    def create_window(self, width_window, height_window):
        """ Create main window. """
        self.window = tkinter.Tk()
        self.window.title('Контроль расходов')
        self.window['bg'] = '#999'
        self.display_width = self.window.winfo_screenwidth()
        self.display_height = self.window.winfo_screenheight()
        self.left = int(self.display_width / 2 - width_window / 2)
        self.top = int(self.display_height / 2 - height_window / 2)
        self.window.geometry(f'{width_window}x{height_window}'
                             f'+{self.left}+{self.top}')

        self.menu = tkinter.Menu(self.window)
        self.window.configure(menu=self.menu)
        self.menu.add_command(label='Добавить',
                              command=lambda: self.append_item(),
                              activebackground='white')
        self.menu.add_command(label='Обновить',
                              command=lambda: self.create_table(),
                              activebackground='white')
        self.menu.add_command(label='Удалить',
                              command=lambda: self.delete_item(),
                              activebackground='white')
        self.create_table()
        self.window.mainloop()

    def create_table(self):
        """ Create table. """
        if self.table:
            self.table.destroy()
        self.table = tkinter.ttk.Treeview(self.window,
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
         for costs, name_costs in change_excel.get_last_costs()]
        self.table.insert(parent='',
                          index=0,
                          values=('', '',
                                  change_excel.summ_value_a_week(
                                      change_excel.get_last_costs())))

        self.window.update()
        self.table.pack(side=tkinter.LEFT, fill='both', expand=True)
        if self.scrollbar:
            self.scrollbar.destroy()
        self.scrollbar = tkinter.ttk.Scrollbar(orient="vertical",
                                               command=self.table.yview)
        self.scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.table['yscrollcommand'] = self.scrollbar.set

    def append_item(self):
        """ Add item into the table. """
        self.append_window = tkinter.Tk()
        self.append_window.title('Добавить')
        self.append_window['bg'] = '#999'
        width_window = 350
        height_window = 110
        self.append_window.geometry(f'{width_window}x{height_window}'
                                    f'+{self.left}+{self.top}')
        self.append_window.resizable(width=False, height=False)

        self.frame_top = tkinter.ttk.Frame(self.append_window,
                                           width=300,
                                           height=45,
                                           borderwidth=10)
        self.frame_bottom = tkinter.ttk.Frame(self.append_window,
                                              width=300,
                                              height=45,
                                              borderwidth=10)

        self.check = (self.append_window.register(self.is_valid), "%P")
        self.cost_value = tkinter.ttk.Entry(master=self.frame_top,
                                            validate="key",
                                            validatecommand=self.check)
        self.cost_name = tkinter.ttk.Entry(master=self.frame_bottom)
        self.value_label = tkinter.ttk.Label(master=self.frame_top,
                                             text='Сколько потрачено:')
        self.name_label = tkinter.ttk.Label(master=self.frame_bottom,
                                            text='На что потрачено:')
        self.add_button = (tkinter.
                           ttk.Button(master=self.append_window,
                                      text='Добавить',
                                      command=lambda:
                                      (change_excel.write_data([
                                          ('' if self.cost_value.get() == ''
                                           else int(self.cost_value.get()),
                                           self.cost_name.get())],
                                          change_excel.get_last_costs()),
                                       self.view_show_message()),
                                      state=tkinter.DISABLED))

        self.value_label.pack(side=tkinter.LEFT)
        self.name_label.pack(side=tkinter.LEFT)
        self.frame_top.pack(fill='both')
        self.frame_bottom.pack(fill='both')
        self.cost_value.pack(side=tkinter.RIGHT)
        self.cost_name.pack(side=tkinter.RIGHT)
        self.add_button.pack(fill='both')
        self.append_window.mainloop()

    def view_show_message(self):
        """ View show message. """
        self.info_message = (
                messagebox.showinfo('', 'Значение добавлено!'))
        self.create_table()
        self.destroy(self.append_window)

    def is_valid(self, check_value):
        """ Check valid data. """
        result = match("[0-9]", check_value) is not None
        if not result:
            self.add_button["state"] = tkinter.DISABLED
        else:
            self.add_button["state"] = tkinter.ACTIVE
        return result

    def delete_item(self):
        """ Delete item which focus. """
        change_excel.delete_data(self.table.selection())
        for item in self.table.selection():
            self.table.delete(item)

    def destroy(self, name_window=None):
        """ Destroy window. """
        if name_window is None:
            name_window = self.window
        name_window.destroy()


if __name__ == '__main__':
    change_excel.write_data(
        change_excel.get_data_from_email(
            get_messages.get_messages_from_email()),
        change_excel.get_last_costs())

    change_excel.delete_data_in_the_weekend(change_excel.get_last_costs())

    window = Window()
    window.create_window(800, 300)
