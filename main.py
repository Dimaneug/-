import pyodbc
from tkinter import ttk
from tkinter import *
from tkinter.messagebox import showerror


class MyDataBase:

    def __init__(self, window):

        self.wind = window
        self.wind.title('База данных')
        self.table_name = None
        self.headers = None
        self.conn = None
        self.cursor = None

        # Поля для main_window
        self.lbl_table_name = Label(text="Введите название таблицы")
        self.ent_table_name = Entry(width=30)
        self.btn_submit_table_name = Button(text="Подтвердить", command=self.show_table_window)
        self.btn_exit = Button(text="Выход", command=self.exit)

        # Поля для table_window
        self.tree = ttk.Treeview()
        self.btn_insert = Button(text="Вставить", command=self.show_insert_window)
        self.btn_update = Button(text="Изменить", command=self.show_update_window)
        self.btn_delete = Button(text="Удалить", command=self.show_delete_window)
        self.btn_back = Button(text="Вернуться", command=self.back_to_main)

        # Поля для insert_window
        self.frm_insert = Frame()
        self.btn_submit_insert = Button(master=self.frm_insert, text="Подтвердить", command=self.submit_insert)
        self.btn_cancel_insert = Button(master=self.frm_insert, text="Отмена", command=self.back_from_insert)

        # Поля для update_window
        self.update_values = []
        self.frm_update = Frame()
        self.btn_submit_update = Button(master=self.frm_update, text="Подтвердить", command=self.submit_update)
        self.btn_cancel_update = Button(master=self.frm_update, text="Отмена", command=self.back_from_update)

        # Поля для delete_window
        self.delete_values = []
        self.frm_delete = Frame()
        self.btn_submit_delete = Button(master=self.frm_delete, text="Подтвердить", command=self.submit_delete)
        self.btn_cancel_delete = Button(master=self.frm_delete, text="Отмена", command=self.back_from_delete)

        self.lbl_header_list = []
        self.ent_list = []

        self.connect_to_table()
        self.main_window()

    def insert_operation(self):
        parameters = []
        for ent in self.ent_list:
            parameters.append(ent.get())
        parameters = tuple(parameters)
        if parameters[0] == '':
            temp_str = '(?'  # temp_str нужна просто как строка с '?'
            for _ in range(len(parameters) - 2):
                temp_str += ',?'
            temp_str += ')'
            headers_str = ' (' + self.headers[1]
            for header in self.headers[2:]:
                headers_str += ', ' + header
            headers_str += ')'
            print(headers_str)
            query = 'INSERT INTO ' + self.table_name + headers_str + " VALUES " + temp_str
            print(query)
            print(parameters)
            self.run_query(query, parameters[1:])
            self.conn.commit()
        else:
            temp_str = '(?'  # temp_str нужна просто как строка с '?'
            for _ in range(len(parameters)-1):
                temp_str += ', ?'
            temp_str += ')'
            headers_str = ' (' + self.headers[0]  # headers_str нужна как строка со столбцами для запроса
            for header in self.headers[1:]:
                headers_str += ', ' + header
            headers_str += ')'
            print(headers_str)
            query = 'INSERT INTO ' + self.table_name + headers_str + " VALUES " + temp_str
            self.run_query(query, parameters)
            self.conn.commit()

    def update_operation(self):
        parameters = {}
        for i, ent in enumerate(self.ent_list):
            if i == 0:
                continue
            if ent.get() != self.update_values[i]:
                parameters[self.headers[i]] = ent.get()

        temp_str = ''
        for key, value in parameters.items():
            temp_str += key + " = '" + value + "', "
        temp_str = temp_str[:-2] + " "
        query = 'UPDATE ' + self.table_name + " SET " + temp_str + " WHERE " + self.headers[0] + " = " + \
                self.update_values[0]
        self.run_query(query)
        self.conn.commit()

    def delete_operation(self):
        query = 'DELETE FROM ' + self.table_name + " WHERE " + self.headers[0] + " = " + self.delete_values[0]
        self.run_query(query)
        self.conn.commit()


    # Скрывает таблицу и кнопки
    def hide_table(self):
        self.tree.grid_forget()
        self.btn_insert.grid_forget()
        self.btn_update.grid_forget()
        self.btn_delete.grid_forget()
        self.btn_back.grid_forget()

    # Функции для кнопок
    def show_table_window(self):
        self.table_name = self.ent_table_name.get()
        self.lbl_table_name.grid_forget()
        self.ent_table_name.grid_forget()
        self.btn_submit_table_name.grid_forget()
        self.btn_exit.grid_forget()
        self.table_window()

    def exit(self):
        self.wind.quit()

    def back_to_main(self):
        records = self.tree.get_children()
        for element in records:
            self.tree.delete(element)
        self.hide_table()

        self.headers = None
        self.main_window()

    def show_insert_window(self):
        self.hide_table()
        self.insert_window()

    def back_from_insert(self):
        for lbl in self.lbl_header_list:
            lbl.grid_forget()
        self.lbl_header_list.clear()

        for ent in self.ent_list:
            ent.grid_forget()
        self.ent_list.clear()

        self.frm_insert.grid_forget()
        self.table_window()

    def submit_insert(self):
        self.insert_operation()
        self.back_from_insert()

    def show_update_window(self):
        selected = self.tree.focus()  # Выбирается выделенная строка
        if not selected:
            return
        self.hide_table()
        self.update_values = [str(self.tree.item(selected, 'text'))]
        self.update_values += list(self.tree.item(selected, 'values'))
        self.update_window()

    def back_from_update(self):
        for lbl in self.lbl_header_list:
            lbl.grid_forget()
        self.lbl_header_list.clear()

        for ent in self.ent_list:
            ent.grid_forget()
        self.ent_list.clear()

        self.update_values.clear()

        self.frm_update.grid_forget()
        self.table_window()

    def submit_update(self):
        self.update_operation()
        self.back_from_update()

    def show_delete_window(self):
        selected = self.tree.focus()  # Выбирается выделенная строка
        if not selected:
            return
        self.hide_table()
        self.delete_values = [str(self.tree.item(selected, 'text'))]
        self.delete_values += list(self.tree.item(selected, 'values'))
        self.delete_window()

    def back_from_delete(self):
        for lbl in self.lbl_header_list:
            lbl.grid_forget()
        self.lbl_header_list.clear()

        for ent in self.ent_list:
            ent.grid_forget()
        self.ent_list.clear()

        self.delete_values.clear()

        self.frm_delete.grid_forget()
        self.table_window()

    def submit_delete(self):
        self.delete_operation()
        self.back_from_delete()



    # Функции, представляющие собой разные окна приложения
    def main_window(self):
        self.lbl_table_name.grid(row=0, column=0, columnspan=2, pady=(20, 5))
        self.ent_table_name.grid(row=1, column=0, columnspan=2)
        self.btn_submit_table_name.grid(row=2, column=0, columnspan=2)
        self.btn_exit.grid(row=3, column=1, pady=10, padx=5, sticky="e")

    def run_query(self, query, parameters=()):
        try:
            self.cursor = self.conn.cursor()
            self.cursor.execute(query, parameters)
        except pyodbc.Error as e:
            self.show_error('run_query', e)

    def table_window(self):
        query = "SELECT * FROM " + self.table_name
        self.run_query(query)
        # Заполняем заголовки столбцов
        self.headers = tuple([i[0] for i in self.cursor.description])
        self.tree = ttk.Treeview(height=10, columns=self.headers[:-1])
        for i, header in enumerate(self.headers, 0):
            self.tree.heading('#'+str(i), text=header, anchor='w')
            self.tree.column('#'+str(i), width=100, stretch=NO)

        # Добавляем все строки таблицы
        for row in self.cursor.fetchall():
            self.tree.insert('', 0, text=row[0], values=tuple([row[i+1] for i in range(len(self.headers)-1)]))
        self.tree.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.btn_insert.grid(row=0, column=1)
        self.btn_update.grid(row=1, column=1)
        self.btn_delete.grid(row=2, column=1)
        self.btn_back.grid(row=3, column=1)

    def insert_window(self):
        for i, header in enumerate(self.headers):
            lbl = Label(text=header, width=20)
            self.lbl_header_list.append(lbl)
            lbl.grid(row=0, column=i)

        for i in range(len(self.headers)):
            ent = Entry(width=20)
            self.ent_list.append(ent)
            ent.grid(row=1, column=i)

        self.btn_submit_insert.grid(row=0, column=0, sticky='e', padx=5)
        self.btn_cancel_insert.grid(row=0, column=1, sticky='w', padx=5)
        self.frm_insert.grid(row=2, column=0, columnspan=len(self.headers))

    def update_window(self):
        for i, header in enumerate(self.headers):
            lbl = Label(text=header, width=20)
            self.lbl_header_list.append(lbl)
            lbl.grid(row=0, column=i)

        for i in range(len(self.headers)):
            if i == 0:
                ent = Label(width=20, text=self.update_values[i])
            else:
                ent = Entry(width=20)
                ent.insert(0, self.update_values[i])
            self.ent_list.append(ent)
            ent.grid(row=1, column=i)

        self.btn_submit_update.grid(row=0, column=0, sticky='e', padx=5)
        self.btn_cancel_update.grid(row=0, column=1, sticky='w', padx=5)
        self.frm_update.grid(row=2, column=0, columnspan=len(self.headers))

    def delete_window(self):
        for i, header in enumerate(self.headers):
            lbl = Label(text=header, width=20)
            self.lbl_header_list.append(lbl)
            lbl.grid(row=0, column=i)

        for i in range(len(self.headers)):
            ent = Label(width=20, text=self.delete_values[i])
            self.ent_list.append(ent)
            ent.grid(row=1, column=i)

        self.btn_submit_delete.grid(row=0, column=0, sticky='e', padx=5)
        self.btn_cancel_delete.grid(row=0, column=1, sticky='w', padx=5)
        self.frm_delete.grid(row=2, column=0, columnspan=len(self.headers))

    # Подключение к базе данных
    def connect_to_table(self):
        try:
            con_string = r"DRIVER={Microsoft Access Driver (*.mdb, " \
                         r"*.accdb)};DBQ=.\bd7.accdb; "
            self.conn = pyodbc.connect(con_string)
            print("Connected To Database")

        except pyodbc.Error as e:
            self.show_error('connect_to_table', e)

    def show_error(self, where, error):
        showerror(title="Ошибка в "+where, message=error)


def main():
    root = Tk()
    application = MyDataBase(root)
    root.mainloop()


if __name__ == "__main__":
    main()
