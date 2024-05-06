# Добовление необходимых библиотек
import os
import sqlite3db
from tkinter import filedialog
from openpyxl import load_workbook
from docxtpl import DocxTemplate
from tkinter import *
from tkinter import ttk
import tkinter as tk
import sqlite3


class GrateWindow:
    def __init__(self, window):
        # Создание GUI приложения на Tkinter
        self.typeCh = None
        self.clientCh = None
        self.selection1 = None
        self.selection2 = None
        self.tree = None
        self.window = window
        window.title("Добро пожаловать в приложение maXim")
        window.geometry('1900x940+0+50')
        window.configure(background="#B2DFDB", )
        # window.attributes("-fullscreen", True)
        self.frame2 = tk.Frame(self.window, width=1900, relief=tk.SUNKEN, borderwidth=1)
        self.frame2.pack(fill=tk.BOTH, side=tk.BOTTOM, expand=True, padx=6, pady=6)
        self.frame1 = tk.Frame(self.window, width=1600, relief=tk.SUNKEN, borderwidth=1, background="#B2DFDB")
        self.frame1.pack(fill=tk.BOTH, side=tk.LEFT, expand=True, padx=6, pady=3)
        self.frame3 = tk.Frame(self.window, width=300, relief=tk.SUNKEN, borderwidth=1)
        self.frame3.pack(fill=tk.BOTH, side=tk.LEFT, expand=True, padx=6, pady=6)
        # Многострочное тектовое окно
        self.l1 = Label(self.frame3, text="Добавить комментарий к договору")
        self.l1.pack(side=tk.TOP)
        self.t1 = tk.Text(self.frame3, width=40, height=19)
        self.t1.pack(side=tk.BOTTOM)
        # Отображение элементов на окне frame1
        lbl0 = Label(self.frame1, text="Заполните все поля для загрузки в форму", font=("Arial Bold", 12), background="#B2DFDB")
        lbl0.grid(column=3, row=1)
        lbl1 = Label(self.frame1, text='Тип договора', background="#B2DFDB")
        lbl1.grid(column=2, row=2)
        lbl2 = Label(self.frame1, text='Поставщик', background="#B2DFDB")
        lbl2.grid(column=2, row=3)
        lbl3 = Label(self.frame1, text='Поставщик(полное)', background="#B2DFDB")
        lbl3.grid(column=2, row=4)
        lbl4 = Label(self.frame1, text='Руководитель(должность)', background="#B2DFDB")
        lbl4.grid(column=2, row=5)
        lbl5 = Label(self.frame1, text='ФИО(инициалы)', background="#B2DFDB")
        lbl5.grid(column=2, row=6)
        lbl6 = Label(self.frame1, text='ФИО(полное)', background="#B2DFDB")
        lbl6.grid(column=2, row=7)
        lbl7 = Label(self.frame1, text='Реквизиты', background="#B2DFDB")
        lbl7.grid(column=2, row=8)
        lbl8 = Label(self.frame1, text='Договор поставки №', background="#B2DFDB")
        lbl8.grid(column=2, row=9)
        lbl9 = Label(self.frame1, text='Адрес объекта', background="#B2DFDB")
        lbl9.grid(column=2, row=10)
        lbl10 = Label(self.frame1, text='Объект(полное)', background="#B2DFDB")
        lbl10.grid(column=2, row=11)
        lbl11 = Label(self.frame1, text='Сумма договора', background="#B2DFDB")
        lbl11.grid(column=2, row=12)
        lbl12 = Label(self.frame1, text='Город', background="#B2DFDB")
        lbl12.grid(column=2, row=13)
        lbl13 = Label(self.frame1, text='Дата', background="#B2DFDB")
        lbl13.grid(column=2, row=14)
        lbl_sqlite = Label(self.frame1, text='Загрузить данные поставщика в базу данных', background="#B2DFDB")
        lbl_sqlite.grid(column=0, row=10)
        lbl_num4 = Label(self.frame1, text='Тип договора', background="#B2DFDB")
        lbl_num4.grid(column=0, row=4)
        lbl_num4 = Label(self.frame1, text='Выбор поставщика', background="#B2DFDB")
        lbl_num4.grid(column=0, row=6)
        lbl_sum = Label(self.frame1, text='Общяя сумма контрактов', background="#B2DFDB")
        lbl_sum.grid(column=1, row=13)
        lbl_number = Label(self.frame1, text='Поиск по номеру договора', background="#B2DFDB")
        lbl_number.grid(column=0, row=8)
        # # Виджеты для ввода текста
        self.txt1 = Entry(self.frame1, width=140)
        self.txt1.grid(column=3, row=2)
        self.txt2 = Entry(self.frame1, width=140)
        self.txt2.grid(column=3, row=3)
        self.txt3 = Entry(self.frame1, width=140)
        self.txt3.grid(column=3, row=4)
        self.txt4 = Entry(self.frame1, width=140)
        self.txt4.grid(column=3, row=5)
        self.txt5 = Entry(self.frame1, width=140)
        self.txt5.grid(column=3, row=6)
        self.txt6 = Entry(self.frame1, width=140)
        self.txt6.grid(column=3, row=7)
        self.txt7 = Entry(self.frame1, width=140)
        self.txt7.grid(column=3, row=8)
        self.txt8 = Entry(self.frame1, width=140)
        self.txt8.grid(column=3, row=9)
        self.txt9 = Entry(self.frame1, width=140)
        self.txt9.grid(column=3, row=10)
        self.txt10 = Entry(self.frame1, width=140)
        self.txt10.grid(column=3, row=11)
        self.txt11 = Entry(self.frame1, width=140)
        self.txt11.grid(column=3, row=12)
        self.txt12 = Entry(self.frame1, width=140)
        self.txt12.grid(column=3, row=13)
        self.txt13 = Entry(self.frame1, width=140)
        self.txt13.grid(column=3, row=14)
        self.txt_search = Entry(self.frame1, width=30)
        self.txt_search.grid(column=0, row=2)
        self.txt_sum = Entry(self.frame1, width=30)
        self.txt_sum.grid(column=1, row=14)
        self.txt_number = Entry(self.frame1, width=30)
        self.txt_number.grid(column=0, row=9)
        btn_sqlite = Button(self.frame1, text="Добавить новый договор", command=self.dataSqlite, width=30)
        btn_sqlite.grid(column=0, row=11)
        btn3 = Button(self.frame1, text="Закрыть приложение", command=self.close_window, width=30)
        btn3.grid(column=1, row=1)
        b1 = tk.Button(self.frame1, text="   Вывести на экран    ", command=self.secondWindow, width=30)
        b1.grid(column=0, row=14)
        btn_to_xlsx = Button(self.frame1, text="На печать (отчет)", command=self.saveExcel, width=30)
        btn_to_xlsx.grid(column=1, row=1)
        btn_search = tk.Button(self.frame1, text="Поиск и заполнить поля формы", command=self.searchDb, width=30)
        btn_search.grid(column=0, row=3)
        btn_update = tk.Button(self.frame1, text="Обновить договор", command=self.updateData, width=30)
        btn_update.grid(column=0, row=13)
        btn_clear = tk.Button(self.frame1, text="Очистить все", command=self.clearText, width=30)
        btn_clear.grid(column=2, row=1)
        btn_sql_3 = Button(self.frame1, text="На печать (форму)", command=self.workSql, width=30)
        btn_sql_3.grid(column=0, row=1)
        self.scroll()

        # Отображение элементов на окне frame2
        self.tree = ttk.Treeview(self.frame2,
                                 columns=("c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8",
                                          "c9", "c10", "c11", "c12", "c13"), show='headings')
        self.tree.text = Entry()
        # self.tree.grid(column=0, row=0)
        self.tree.column("#1", anchor=tk.CENTER)
        self.tree.heading("#1", text="Тип договора")
        self.tree.column("#2", anchor=tk.CENTER)
        self.tree.heading("#2", text="Поставщик")
        self.tree.column("#3", anchor=tk.CENTER)
        self.tree.heading("#3", text="Поставщик(полное)")
        self.tree.column("#4", anchor=tk.CENTER)
        self.tree.heading("#4", text="Должность")
        self.tree.column("#5", anchor=tk.CENTER)
        self.tree.heading("#5", text="ФИО")
        self.tree.column("#6", anchor=tk.CENTER)
        self.tree.heading("#6", text="ФИО(полное)")
        self.tree.column("#7", anchor=tk.CENTER)
        self.tree.heading("#7", text="Реквизиты")
        self.tree.column("#8", anchor=tk.CENTER)
        self.tree.heading("#8", text="Номер договора")
        self.tree.column("#9", anchor=tk.CENTER)
        self.tree.heading("#9", text="Объект")
        self.tree.column("#10", anchor=tk.CENTER)
        self.tree.heading("#10", text="Объект(полное)")
        self.tree.column("#11", anchor=tk.CENTER)
        self.tree.heading("#11", text="Сумма договора")
        self.tree.column("#12", anchor=tk.CENTER)
        self.tree.heading("#12", text="Город")
        self.tree.column("#13", anchor=tk.CENTER)
        self.tree.heading("#13", text="Дата")
        self.tree.pack(side=LEFT)
        for i in self.tree['columns']:
            self.tree.column(i, minwidth=50, width=145, stretch=NO)
        self.tree.config(height=28)
        style = ttk.Style()
        style.theme_use("clam")
        style.map("Treeview")
        # Прокрутка таблицы в frame2
        scroll_y = tk.Scrollbar(self.frame2, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_y.pack(side=RIGHT, expand=True, fill="y")
        self.tree.configure(yscrollcommand=scroll_y.set)

    def selected1(self, event):
        # получаем выделенный элемент
        self.selection1 = str(self.clientCh.get()).replace('{', '').replace('}', '')
        self.txt_search.insert(0, self.selection1)
        self.clientCh.set('')

    def selected2(self, event):
        # получаем выделенный элемент
        self.selection2 = str(self.typeCh.get())
        self.txt1.insert(0, self.selection2)
        self.typeCh.set('')
        
    def scroll(self):
        # Combobox creation
        n2 = tk.StringVar()
        self.clientCh = ttk.Combobox(self.frame1, width=30, textvariable=n2)
        # Adding combobox drop down list
        self.clientCh['values'] = self.scrollDb()
        self.clientCh.grid(column=0, row=7)
        self.clientCh.current()
        self.clientCh.bind("<<ComboboxSelected>>", self.selected1)
        # Adding combobox drop down list
        n3 = tk.StringVar()
        self.typeCh = ttk.Combobox(self.frame1, width=30, textvariable=n3)
        self.typeCh['values'] = ["Договор поставки", "Договор аренды", "Договор работ"]
        self.typeCh.grid(column=0, row=5)
        self.typeCh.current()
        self.typeCh.bind("<<ComboboxSelected>>", self.selected2)

    # Функция, после заполнеия текстовых полей, при нажатии на кнопку, добавление записи в SQLite
    def dataSqlite(self):
        s1 = self.txt1.get()
        s2 = self.txt2.get()
        s3 = self.txt3.get()
        s4 = self.txt4.get()
        s5 = self.txt5.get()
        s6 = self.txt6.get()
        s7 = self.txt7.get()
        s8 = self.txt8.get()
        s9 = self.txt9.get()
        s10 = self.txt10.get()
        s11 = self.txt11.get()
        s12 = self.txt12.get()
        s13 = self.txt13.get()
        s14 = self.t1.get("1.0", tk.END)
        if not s8:
            print("Empty Variable")
        else:
            sqlite3db.insert_varible_into_table(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13)
            sqlite3db.insert_varible_into_table2(s8, s14)
      
    # Кнопка закрытия приложения
    def close_window(self):
        self.window.destroy()
    
    # # Функция, выводит на экран таблицу базы данных в отдельном окне
    def secondWindow(self):
        # second = Toplevel(window)
        # second.title("База данных")
        self.txt_sum.delete(0, END)
        # Перед заполнение таблицы данными, удаляем предыдущие
        for item in self.tree.get_children():
            self.tree.delete(item)
        x_client1 = self.txt_search.get()
        y_typedoc = self.txt1.get()
        if not x_client1 and not y_typedoc:
            con1 = sqlite3.connect("sqlite_python.db")
            cur1 = con1.cursor()
            cur1.execute('''SELECT * FROM ProgramFin''')
            rows = cur1.fetchall()
            con2 = sqlite3.connect("sqlite_python.db")
            cur2 = con2.cursor()
            cur2.execute('''SELECT SUM(price) FROM ProgramFin''')
            rows2 = cur2.fetchall()
            for row1 in rows:
                # print(row1)
                self.tree.insert("", tk.END, values=row1)
            for row2 in rows2:
                # print(row2)
                self.txt_sum.insert(0, row2)
                self.tree.insert("", tk.END, values=row2)
            con1.close()
            con2.close()
        else:
            if not x_client1 or not y_typedoc:
                con1 = sqlite3.connect("sqlite_python.db")
                cur1 = con1.cursor()
                cur1.execute(f'''SELECT * FROM ProgramFin WHERE idArt=? OR client1=?''', (y_typedoc, x_client1,))
                rows = cur1.fetchall()
                con2 = sqlite3.connect("sqlite_python.db")
                cur2 = con2.cursor()
                cur2.execute(f'''SELECT SUM(price) FROM ProgramFin WHERE idArt=? OR client1=?''', (y_typedoc, x_client1,))
                rows2 = cur2.fetchall()
                for row1 in rows:
                    # print(row1)
                    self.tree.insert("", tk.END, values=row1)
                for row2 in rows2:
                    # print(row2)
                    self.txt_sum.insert(0, row2)
                    self.tree.insert("", tk.END, values=row2)
                con1.close()
                con2.close()
            else:
                con1 = sqlite3.connect("sqlite_python.db")
                cur1 = con1.cursor()
                cur1.execute(f'''SELECT * FROM ProgramFin WHERE idArt=? AND client1=?''', (y_typedoc, x_client1,))
                rows = cur1.fetchall()
                con2 = sqlite3.connect("sqlite_python.db")
                cur2 = con2.cursor()
                cur2.execute(f'''SELECT SUM(price) FROM ProgramFin WHERE idArt=? AND client1=?''', (y_typedoc, x_client1,))
                rows2 = cur2.fetchall()
                for row1 in rows:
                    # print(row1)
                    self.tree.insert("", tk.END, values=row1)
                for row2 in rows2:
                    # print(row2)
                    self.txt_sum.insert(0, row2)
                    self.tree.insert("", tk.END, values=row2)
                con1.close()
                con2.close()
        self.txt_search.delete(0, END)
        self.txt1.delete(0, END)

    # Сохранение файла xlsx и вывод на печать
    def saveExcel(self):
        workbook = load_workbook(filename='XLSX_TO_PRINT/Table1.xlsx')
        sheet = workbook['Sheet1']
        sheet.delete_rows(idx=2, amount=15)
        for row_id in self.tree.get_children():
            row = self.tree.item(row_id)['values']
            sheet.append(row)
        workbook.save(filename='XLSX_TO_PRINT/Table1.xlsx')

        file_path = filedialog.askopenfilename(title="Файл", initialdir="XLSX_TO_PRINT/",
                                               filetypes=[("Text File", '*.xlsx'), ("All files", "*.*")])
        print("Selected File:", file_path)
        os.startfile(file_path)

    # # Функция, после точного ввода имени поставщика, поиск в БД столбца и вывод строки заполением текстовых полей
    def searchDb(self):
        d = self.txt_number.get()
        conn = sqlite3.connect('sqlite_python.db')
        cursor1 = conn.cursor()
        cursor1.execute(f'''SELECT * FROM ProgramFin
                        WHERE dog=?''', (d,))
        rows = cursor1.fetchone()
        clist1 = []
        for row in rows:
            clist1.append(row)
        # print(clist1)
        # print(len(clist1))
        self.txt1.insert(0, clist1[0])
        self.txt2.insert(0, clist1[1])
        self.txt3.insert(0, clist1[2])
        self.txt4.insert(0, clist1[3])
        self.txt5.insert(0, clist1[4])
        self.txt6.insert(0, clist1[5])
        self.txt7.insert(0, clist1[6])
        self.txt8.insert(0, clist1[7])
        self.txt9.insert(0, clist1[8])
        self.txt10.insert(0, clist1[9])
        self.txt11.insert(0, clist1[10])
        self.txt12.insert(0, clist1[11])
        self.txt13.insert(0, clist1[12])
        # Вывод комментариев из второй таблицы
        cursor2 = conn.cursor()
        cursor2.execute(f'''SELECT * FROM ProgramFin2
                        WHERE dog=?''', (d,))
        rows = cursor2.fetchone()
        clist2 = []
        for row in rows:
            clist2.append(row)

        self.t1.insert("1.0", clist2[1])
        conn.close()
        self.txt_search.delete(0, END)

    # Обновление строки по ключу: номер договора
    def updateData(self):
        s1 = self.txt1.get()
        s2 = self.txt2.get()
        s3 = self.txt3.get()
        s4 = self.txt4.get()
        s5 = self.txt5.get()
        s6 = self.txt6.get()
        s7 = self.txt7.get()
        s8 = self.txt8.get()
        s9 = self.txt9.get()
        s10 = self.txt10.get()
        s11 = self.txt11.get()
        s12 = self.txt12.get()
        s13 = self.txt13.get()
        s14 = self.t1.get("1.0", tk.END)
        conn = sqlite3.connect('sqlite_python.db')
        cursor1 = conn.cursor()
        cursor1.execute(f'''UPDATE ProgramFin SET idArt='{s1}', client1='{s2}', ful_name='{s3}', boss='{s4}', name='{s5}',
                            face='{s6}', client2='{s7}', object2='{s9}', object1='{s10}', price='{s11}', city='{s12}',
                            time='{s13}' WHERE dog = '{s8}' ''')
        # Добавление комментариев во вторую таблицу
        cursor2 = conn.cursor()
        cursor2.execute(f'''UPDATE ProgramFin2 SET commitM='{s14}' WHERE dog='{s8}' ''')
        conn.commit()
        conn.close()

    # Очищает текстовые поля приложения
    def clearText(self):
        self.txt1.delete(0, END)
        self.txt2.delete(0, END)
        self.txt3.delete(0, END)
        self.txt4.delete(0, END)
        self.txt5.delete(0, END)
        self.txt6.delete(0, END)
        self.txt7.delete(0, END)
        self.txt8.delete(0, END)
        self.txt9.delete(0, END)
        self.txt10.delete(0, END)
        self.txt11.delete(0, END)
        self.txt12.delete(0, END)
        self.txt13.delete(0, END)
        self.txt_sum.delete(0, END)
        self.txt_search.delete(0, END)
        self.txt_number.delete(0, END)
        self.t1.delete("1.0", END)

    # Функция, запрос имени поставщика, заполняет все данные по нему из текстовых полей в формы договоров и на печать
    def workSql(self):
        typedoc = self.txt1.get()
        client1 = self.txt2.get()
        ful_name = self.txt3.get()
        boss = self.txt4.get()
        name = self.txt5.get()
        face = self.txt6.get()
        client2 = self.txt7.get()
        dog = self.txt8.get()
        object2 = self.txt9.get()
        object1 = self.txt10.get()
        price = self.txt11.get()
        city = self.txt12.get()
        time = self.txt13.get()

        context = {'client1': client1,
                   'ful_name': ful_name,
                   'boss': boss,
                   'name': name,
                   'face': face,
                   'client2': client2,
                   'dog': dog,
                   'object2': object2,
                   'object1': object1,
                   'price': price,
                   'city': city,
                   "time": time
                   }

        if typedoc == "Договор поставки":
            doc = DocxTemplate('Форма_договора_поставки.docx')
            doc.render(context)
            doc.save("DOC_TO_PRINT/Договор_поставки_на_печать.docx")
            file_path = filedialog.askopenfilename(title="Файл", initialdir="DOC_TO_PRINT/",
                                                   filetypes=[("Text File", '*.docx'), ("All files", "*.*")])
            print("Selected File:", file_path)
            os.startfile(file_path)

        elif typedoc == "Договор аренды":
            doc = DocxTemplate('Форма_договора_аренды.docx')
            doc.render(context)
            doc.save("DOC_TO_PRINT/Договор_аренды_на_печать.docx")
            file_path = filedialog.askopenfilename(title="Файл", initialdir="DOC_TO_PRINT/",
                                                   filetypes=[("Text File", '*.docx'), ("All files", "*.*")])
            print("Selected File:", file_path)
            os.startfile(file_path)

        elif typedoc == "Договор работ":
            doc = DocxTemplate('Форма_договора_работ.docx')
            doc.render(context)
            doc.save("DOC_TO_PRINT/Договор_работ_на_печать.docx")
            file_path = filedialog.askopenfilename(title="Файл", initialdir="DOC_TO_PRINT/",
                                                   filetypes=[("Text File", '*.docx'), ("All files", "*.*")])
            print("Selected File:", file_path)
            os.startfile(file_path)

    # Функция, после точного ввода имени поставщика, поиск в БД столбца и вывод строки заполением текстовых полей
    def scrollDb(self):
        conn = sqlite3.connect('sqlite_python.db')
        c = conn.cursor()
        c.execute('''SELECT client1 FROM ProgramFin''')
        rows = c.fetchall()
        clist1 = []
        for row in rows:
            clist1.append(row)
        # print(clist1)
        # print(len(clist1))
        conn.close()
        return list(set(clist1))  # способ, с помощью которого дубликаты удаляются из списка


def main():
    window = tk.Tk()
    GrateWindow(window)
    window.mainloop()


if __name__ == '__main__':
    main()

