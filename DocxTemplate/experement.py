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

# Создание GUI приложения на Tkinter
window = tk.Tk()
window.title("Добро пожаловать в приложение maXim")
window.geometry('1900x940+0+50')
window.configure(background="#B2DFDB", )
# window.attributes("-fullscreen", True)


frame2 = tk.Frame(window, width=1900, relief=tk.SUNKEN, borderwidth=1)
frame2.pack(fill=tk.BOTH, side=tk.BOTTOM, expand=True, padx=6, pady=6)
frame1 = tk.Frame(window, width=1600, relief=tk.SUNKEN, borderwidth=1, background="#B2DFDB")
frame1.pack(fill=tk.BOTH, side=tk.LEFT, expand=True, padx=6, pady=3)
frame3 = tk.Frame(window, width=300, relief=tk.SUNKEN, borderwidth=1)
frame3.pack(fill=tk.BOTH, side=tk.LEFT, expand=True, padx=6, pady=6)
frame4 = tk.Frame(window, width=1900, relief=tk.SUNKEN, borderwidth=1)
frame4.pack(fill=tk.BOTH, side=tk.BOTTOM, expand=True, padx=6, pady=6)

# Многострочное тектовое окно
l1 = Label(frame3, text="Добавить комментарий к договору")
l1.pack(side=tk.TOP)
t1 = tk.Text(frame3, width=40, height=19)
t1.pack(side=tk.BOTTOM)
# Отображение элементов на окне frame1
lbl0 = Label(frame1, text="Заполните все поля для загрузки в форму", font=("Arial Bold", 12), background="#B2DFDB")
lbl0.grid(column=3, row=1)
lbl1 = Label(frame1, text='Тип договора', background="#B2DFDB")
lbl1.grid(column=2, row=2)
lbl2 = Label(frame1, text='Поставщик', background="#B2DFDB")
lbl2.grid(column=2, row=3)
lbl3 = Label(frame1, text='Поставщик(полное)', background="#B2DFDB")
lbl3.grid(column=2, row=4)
lbl4 = Label(frame1, text='Руководитель(должность)', background="#B2DFDB")
lbl4.grid(column=2, row=5)
lbl5 = Label(frame1, text='ФИО(инициалы)', background="#B2DFDB")
lbl5.grid(column=2, row=6)
lbl6 = Label(frame1, text='ФИО(полное)', background="#B2DFDB")
lbl6.grid(column=2, row=7)
lbl7 = Label(frame1, text='Реквизиты', background="#B2DFDB")
lbl7.grid(column=2, row=8)
lbl8 = Label(frame1, text='Договор поставки №', background="#B2DFDB")
lbl8.grid(column=2, row=9)
lbl9 = Label(frame1, text='Адрес объекта', background="#B2DFDB")
lbl9.grid(column=2, row=10)
lbl10 = Label(frame1, text='Объект(полное)', background="#B2DFDB")
lbl10.grid(column=2, row=11)
lbl11 = Label(frame1, text='Сумма договора', background="#B2DFDB")
lbl11.grid(column=2, row=12)
lbl12 = Label(frame1, text='Город', background="#B2DFDB")
lbl12.grid(column=2, row=13)
lbl13 = Label(frame1, text='Дата', background="#B2DFDB")
lbl13.grid(column=2, row=14)
lbl_sqlite = Label(frame1, text='Загрузить данные поставщика в базу данных', background="#B2DFDB")
lbl_sqlite.grid(column=0, row=10)
lbl_num4 = Label(frame1, text='Тип договора', background="#B2DFDB")
lbl_num4.grid(column=0, row=4)
lbl_num4 = Label(frame1, text='Выбор поставщика', background="#B2DFDB")
lbl_num4.grid(column=0, row=6)
lbl_sum = Label(frame1, text='Общяя сумма контрактов', background="#B2DFDB")
lbl_sum.grid(column=1, row=13)
lbl_number = Label(frame1, text='Поиск по номеру договора', background="#B2DFDB")
lbl_number.grid(column=0, row=8)
# # Виджеты для ввода текста
txt1 = Entry(frame1, width=140)
txt1.grid(column=3, row=2)
txt2 = Entry(frame1, width=140)
txt2.grid(column=3, row=3)
txt3 = Entry(frame1, width=140)
txt3.grid(column=3, row=4)
txt4 = Entry(frame1, width=140)
txt4.grid(column=3, row=5)
txt5 = Entry(frame1, width=140)
txt5.grid(column=3, row=6)
txt6 = Entry(frame1, width=140)
txt6.grid(column=3, row=7)
txt7 = Entry(frame1, width=140)
txt7.grid(column=3, row=8)
txt8 = Entry(frame1, width=140)
txt8.grid(column=3, row=9)
txt9 = Entry(frame1, width=140)
txt9.grid(column=3, row=10)
txt10 = Entry(frame1, width=140)
txt10.grid(column=3, row=11)
txt11 = Entry(frame1, width=140)
txt11.grid(column=3, row=12)
txt12 = Entry(frame1, width=140)
txt12.grid(column=3, row=13)
txt13 = Entry(frame1, width=140)
txt13.grid(column=3, row=14)
txt_search = Entry(frame1, width=30)
txt_search.grid(column=0, row=2)
txt_sum = Entry(frame1, width=30)
txt_sum.grid(column=1, row=14)
txt_number = Entry(frame1, width=30)
txt_number.grid(column=0, row=9)


# Функция, после заполнеия текстовых полей, при нажатии на кнопку, добавление записи в SQLite
def dataSqlite():
    s1 = txt1.get()
    s2 = txt2.get()
    s3 = txt3.get()
    s4 = txt4.get()
    s5 = txt5.get()
    s6 = txt6.get()
    s7 = txt7.get()
    s8 = txt8.get()
    s9 = txt9.get()
    s10 = txt10.get()
    s11 = txt11.get()
    s12 = txt12.get()
    s13 = txt13.get()
    s14 = t1.get("1.0", tk.END)
    if not s8:
        print("Empty Variable")
    else:
        sqlite3db.insert_varible_into_table(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13)
        sqlite3db.insert_varible_into_table2(s8, s14)


btn_sqlite = Button(frame1, text="Добавить новый договор", command=dataSqlite, width=30)
btn_sqlite.grid(column=0, row=11)


# Кнопка закрытия приложения
def close_window():
    window.destroy()


btn3 = Button(frame1, text="Закрыть приложение", command=close_window, width=30)
btn3.grid(column=1, row=1)


# Отображение элементов на окне frame2

def treeWindow():
    tree = ttk.Treeview(frame2, columns=("c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", "c11", "c12", "c13"),
                        show='headings')
    tree.text = Entry()
    tree.grid(column=0, row=0)
    tree.column("#1", anchor=tk.CENTER)
    tree.heading("#1", text="Тип договора")
    tree.column("#2", anchor=tk.CENTER)
    tree.heading("#2", text="Поставщик")
    tree.column("#3", anchor=tk.CENTER)
    tree.heading("#3", text="Поставщик(полное)")
    tree.column("#4", anchor=tk.CENTER)
    tree.heading("#4", text="Должность")
    tree.column("#5", anchor=tk.CENTER)
    tree.heading("#5", text="ФИО")
    tree.column("#6", anchor=tk.CENTER)
    tree.heading("#6", text="ФИО(полное)")
    tree.column("#7", anchor=tk.CENTER)
    tree.heading("#7", text="Реквизиты")
    tree.column("#8", anchor=tk.CENTER)
    tree.heading("#8", text="Номер договора")
    tree.column("#9", anchor=tk.CENTER)
    tree.heading("#9", text="Объект")
    tree.column("#10", anchor=tk.CENTER)
    tree.heading("#10", text="Объект(полное)")
    tree.column("#11", anchor=tk.CENTER)
    tree.heading("#11", text="Сумма договора")
    tree.column("#12", anchor=tk.CENTER)
    tree.heading("#12", text="Город")
    tree.column("#13", anchor=tk.CENTER)
    tree.heading("#13", text="Дата")
    tree.pack(side=LEFT)
    for i in tree['columns']:
        tree.column(i, minwidth=50, width=145, stretch=NO)
    tree.config(height=28)
    # Прокрутка таблицы в frame2
    scroll_y = tk.Scrollbar(frame2, orient=tk.VERTICAL, command=tree.yview)
    scroll_y.pack(side=RIGHT, expand=True, fill="y")
    tree.configure(yscrollcommand=scroll_y.set)

    style = ttk.Style()
    style.theme_use("clam")
    style.map("Treeview")
    secondWindow(tree)


# Функция, выводит на экран таблицу базы данных в отдельном окне
def secondWindow(tree):
    # second = Toplevel(window)
    # second.title("База данных")
    txt_sum.delete(0, END)
    # Перед заполнение таблицы данными, удаляем предыдущие
    for item in tree.get_children():
        tree.delete(item)
    x_client1 = txt_search.get()
    y_typedoc = txt1.get()
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
            tree.insert("", tk.END, values=row1)
        for row2 in rows2:
            # print(row2)
            txt_sum.insert(0, row2)
            tree.insert("", tk.END, values=row2)
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
                tree.insert("", tk.END, values=row1)
            for row2 in rows2:
                # print(row2)
                txt_sum.insert(0, row2)
                tree.insert("", tk.END, values=row2)
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
                tree.insert("", tk.END, values=row1)
            for row2 in rows2:
                # print(row2)
                txt_sum.insert(0, row2)
                tree.insert("", tk.END, values=row2)
            con1.close()
            con2.close()
    txt_search.delete(0, END)
    txt1.delete(0, END)


b1 = tk.Button(frame1, text="   Вывести на экран    ", command=treeWindow, width=30)
b1.grid(column=0, row=14)


# Сохранение файла xlsx и вывод на печать
def saveExcel(tree):
    workbook = load_workbook(filename='XLSX_TO_PRINT/Table1.xlsx')
    sheet = workbook['Sheet1']
    # Удаляем все строки, кроме заголовков
    sheet.delete_rows(2, sheet.max_row - 1)
    workbook.save(filename='XLSX_TO_PRINT/Table1.xlsx')
    # sheet.delete_rows(idx=2, amount=15)
    for row_id in tree.get_children():
        row = tree.item(row_id)['values']
        sheet.append(row)
    workbook.save(filename='XLSX_TO_PRINT/Table1.xlsx')

    file_path = filedialog.askopenfilename(title="Файл", initialdir="XLSX_TO_PRINT/",
                                           filetypes=[("Text File", '*.xlsx'), ("All files", "*.*")])
    print("Selected File:", file_path)
    os.startfile(file_path)


btnSqlToXlsx = Button(frame1, text="На печать (отчет)", command=saveExcel, width=30)
btnSqlToXlsx.grid(column=1, row=1)


def windowNew():
    tree2 = ttk.Treeview(frame2, columns=("c1", "c2", "c3", "c4", "c5"),
                        show='headings')
    tree2.text = Entry()
    tree2.grid(column=0, row=0)
    tree2.column("#1", anchor=tk.CENTER)
    tree2.heading("#1", text="Тип договора")
    tree2.column("#2", anchor=tk.CENTER)
    tree2.heading("#2", text="Поставщик")
    tree2.column("#3", anchor=tk.CENTER)
    tree2.heading("#3", text="Номер договора")
    tree2.column("#4", anchor=tk.CENTER)
    tree2.heading("#4", text="Объект")
    tree2.column("#5", anchor=tk.CENTER)
    tree2.heading("#5", text="Сумма договора")
    tree2.pack(side=LEFT)
    for i in tree2['columns']:
        tree2.column(i, minwidth=50, width=145, stretch=NO)
    tree2.config(height=28)
    # Прокрутка таблицы в frame2
    scroll_y = tk.Scrollbar(frame2, orient=tk.VERTICAL, command=tree.yview)
    scroll_y.pack(side=RIGHT, expand=True, fill="y")
    tree2.configure(yscrollcommand=scroll_y.set)

    style = ttk.Style()
    style.theme_use("default")
    style.map("Treeview")


btn_new = Button(frame1, text="(отчет new)", command=lambda: windowNew, width=30)
btn_new.grid(column=1, row=3)


# # Функция, после точного ввода номера договора, поиск в БД столбца и вывод строки заполением текстовых полей
def searchDb():
    d = txt_number.get()
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
    txt1.insert(0, clist1[0])
    txt2.insert(0, clist1[1])
    txt3.insert(0, clist1[2])
    txt4.insert(0, clist1[3])
    txt5.insert(0, clist1[4])
    txt6.insert(0, clist1[5])
    txt7.insert(0, clist1[6])
    txt8.insert(0, clist1[7])
    txt9.insert(0, clist1[8])
    txt10.insert(0, clist1[9])
    txt11.insert(0, clist1[10])
    txt12.insert(0, clist1[11])
    txt13.insert(0, clist1[12])
    # Вывод комментариев из второй таблицы
    cursor2 = conn.cursor()
    cursor2.execute(f'''SELECT * FROM ProgramFin2
                    WHERE dog=?''', (d,))
    rows = cursor2.fetchone()
    clist2 = []
    for row in rows:
        clist2.append(row)

    t1.insert("1.0", clist2[1])
    conn.close()
    txt_search.delete(0, END)


btn_search = tk.Button(frame1, text="Поиск и заполнить поля формы", command=searchDb, width=30)
btn_search.grid(column=0, row=3)


# Обновление строки по ключу: номер договора
def updateData():
    s1 = txt1.get()
    s2 = txt2.get()
    s3 = txt3.get()
    s4 = txt4.get()
    s5 = txt5.get()
    s6 = txt6.get()
    s7 = txt7.get()
    s8 = txt8.get()
    s9 = txt9.get()
    s10 = txt10.get()
    s11 = txt11.get()
    s12 = txt12.get()
    s13 = txt13.get()
    s14 = t1.get("1.0", tk.END)
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


btn_update = tk.Button(frame1, text="Обновить договор", command=updateData, width=30)
btn_update.grid(column=0, row=13)


# Очищает текстовые поля приложения
def clearText():
    txt1.delete(0, END)
    txt2.delete(0, END)
    txt3.delete(0, END)
    txt4.delete(0, END)
    txt5.delete(0, END)
    txt6.delete(0, END)
    txt7.delete(0, END)
    txt8.delete(0, END)
    txt9.delete(0, END)
    txt10.delete(0, END)
    txt11.delete(0, END)
    txt12.delete(0, END)
    txt13.delete(0, END)
    txt_sum.delete(0, END)
    txt_search.delete(0, END)
    txt_number.delete(0, END)
    t1.delete("1.0", END)


btn_clear = tk.Button(frame1, text="Очистить все", command=clearText, width=30)
btn_clear.grid(column=2, row=1)


# Функция, запрос имени поставщика, заполняет все данные по нему из текстовых полей в формы договоров и на печать
def workSql():
    typedoc = txt1.get()
    client1 = txt2.get()
    ful_name = txt3.get()
    boss = txt4.get()
    name = txt5.get()
    face = txt6.get()
    client2 = txt7.get()
    dog = txt8.get()
    object2 = txt9.get()
    object1 = txt10.get()
    price = txt11.get()
    city = txt12.get()
    time = txt13.get()

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


btnSql3 = Button(frame1, text="На печать (форму)", command=workSql, width=30)
btnSql3.grid(column=0, row=1)


def scrollDb():
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


def selected1(event):
    # получаем выделенный элемент
    selection1 = str(clientCh.get()).replace('{', '').replace('}', '')
    txt_search.insert(0, selection1)
    clientCh.set('')


# Combobox creation
n2 = tk.StringVar()
clientCh = ttk.Combobox(frame1, width=30, textvariable=n2)

# Adding combobox drop down list
clientCh['values'] = scrollDb()
clientCh.grid(column=0, row=7)
clientCh.current()
clientCh.bind("<<ComboboxSelected>>", selected1)


def selected2(event):
    # получаем выделенный элемент
    selection2 = str(typeCh.get())
    txt1.insert(0, selection2)
    typeCh.set('')


n3 = tk.StringVar()
typeCh = ttk.Combobox(frame1, width=30, textvariable=n3)
typeCh['values'] = ["Договор поставки", "Договор аренды", "Договор работ"]
typeCh.grid(column=0, row=5)
typeCh.current()
typeCh.bind("<<ComboboxSelected>>", selected2)


window.mainloop()
