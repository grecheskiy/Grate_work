import sqlite3

sqlite_connection = None

# try:
#     sqlite_connection = sqlite3.connect('sqlite_python.db')
#     sqlite_create_table_query = '''CREATE TABLE ProgramFin (
#                                 idArt INTEGER FOREGIN KEY,
#                                 client1 TEXT NOT NULL,
#                                 ful_name TEXT NOT NULL,
#                                 boss TEXT NOT NULL,
#                                 name TEXT NOT NULL,
#                                 face TEXT NOT NULL,
#                                 client2 TEXT NOT NULL,
#                                 dog TEXT NOT NULL,
#                                 object2 TEXT NOT NULL,
#                                 object1 TEXT NOT NULL,
#                                 price TEXT NOT NULL,
#                                 city TEXT NOT NULL,
#                                 time datetime);'''
#
#     cursor = sqlite_connection.cursor()
#     print("База данных подключена к SQLite")
#     cursor.execute(sqlite_create_table_query)
#     sqlite_connection.commit()
#     print("Таблица SQLite создана")
#
#     cursor.close()
#
# except sqlite3.Error as error:
#     print("Ошибка при подключении к sqlite", error)
# finally:
#     if sqlite_connection:
#         sqlite_connection.close()
#         print("Соединение с SQLite закрыто")


def insert_varible_into_table(idArt, client1, ful_name, boss, name, face, client2, dog, object2, object1, price, city, time):

    global sqlite_connection
    try:
        sqlite_connection = sqlite3.connect('sqlite_python.db', timeout=7)
        cursor = sqlite_connection.cursor()
        print("Подключен к SQLite")

        sqlite_insert_with_param = """INSERT INTO ProgramFin
                                 (idArt, client1, ful_name, boss, name, face, client2, 
                                 dog, object2, object1, price, city, time)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);"""

        data_tuple = (idArt, client1, ful_name, boss, name, face, client2, dog, object2, object1, price, city, time)
        cursor.execute(sqlite_insert_with_param, data_tuple)
        sqlite_connection.commit()
        print("Записи успешно вставлены в таблицу ProgramFin", cursor.rowcount)
        sqlite_connection.commit()
        cursor.close()

    except sqlite3.Error as error:
        print("Ошибка при работе с SQLite", error)
    finally:
        if sqlite_connection:
            sqlite_connection.close()
            print("Соединение с SQLite закрыто")


# insert_varible_into_table(1, 'client1', 'ful_name', 'boss', 'name', 'face', 'client2', 'dog', 'object2', 'object1', 'price', 'city', 'time')

