import sqlite3
import pandas as pd
from pandas import DataFrame


class Database:
    def create_db(self):
        try:
            sqliteConnection = sqlite3.connect('SQLite_Python.db')
            cursor = sqliteConnection.cursor()
            print("Database created and Successfully Connected to SQLite")

            sqlite_select_Query = "select sqlite_version();"
            cursor.execute(sqlite_select_Query)
            record = cursor.fetchall()
            print("SQLite Database Version is: ", record)
            cursor.close()

        except sqlite3.Error as error:
            print("Error while connecting to sqlite", error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                print("The SQLite connection is closed")

    def is_table(self, table_name):
        """ This method seems to be working now"""
        conn = sqlite3.connect('SQLite_Python.db')

        query = "SELECT name from sqlite_master WHERE type='table' AND name='" + table_name + "';"
        cursor = conn.execute(query)
        result = cursor.fetchone()
        if result == None:
            return False
        else:
            return True

    def create_parts1_table(self, name=''):
        try:
            sqliteConnection = sqlite3.connect('SQLite_Python.db')
            sqlite_create_table_query = '''CREATE TABLE {} (
                                        MA_CHI_TIET TEXT PRIMARY KEY,
                                        TEN_CHI_TIET BLOB NOT NULL,
                                        LOAI_VAT_TU	BLOB,
                                        don_vi BLOB,
                                        dai	BLOB,
                                        cao	BLOB,
                                        day	BLOB,
                                        so_luong INTEGER NOT NULL,	
                                        ghi_chu	BLOB,
                                        part_11524 INTEGER,	
                                        part_11523 INTEGER,	
                                        part_11528 INTEGER,	
                                        part_11529 INTEGER,  
                                        part_11549 INTEGER );'''.format(name)

            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")
            cursor.execute(sqlite_create_table_query)
            sqliteConnection.commit()
            print("Parts1 table created")

            cursor.close()

        except sqlite3.Error as error:
            print("Error while creating a sqlite table", error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                print("sqlite connection is closed")

    def create_parts2_table(self, name=''):
        try:
            sqliteConnection = sqlite3.connect('SQLite_Python.db')
            sqlite_create_table_query = '''CREATE TABLE {} (
                                        id TEXT PRIMARY KEY,
                                        name TEXT NOT NULL,
                                        don_vi TEXT,
                                        rong TEXT,
                                        cao TEXT,
                                        sau TEXT,
                                        gia_cong TEXT,
                                        so_luong INTEGER NOT NULL,                                       
                                        ghi_chu TEXT );'''.format(name)

            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")
            cursor.execute(sqlite_create_table_query)
            sqliteConnection.commit()
            print("Parts2 table created")

            cursor.close()

        except sqlite3.Error as error:
            print("Error while creating a sqlite table", error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                print("sqlite connection is closed")

    def create_item_table(self):
        try:
            sqliteConnection = sqlite3.connect('SQLite_Python.db')
            sqlite_create_table_query = '''CREATE TABLE Items (
                                        id INTEGER PRIMARY KEY,
                                        name TEXT NOT NULL,
                                        parts1 TEXT,
                                        parts2 TEXT);'''

            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")
            cursor.execute(sqlite_create_table_query)
            sqliteConnection.commit()
            print("Item table created")

            cursor.close()

        except sqlite3.Error as error:
            print("Error while creating a sqlite table", error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                print("sqlite connection is closed")

    def delete_table(self, table=""):
        # Connecting to sqlite
        conn = sqlite3.connect('SQLite_Python.db')

        # Creating a cursor object using the cursor() method
        cursor = conn.cursor()

        # Doping EMPLOYEE table if already exists
        cursor.execute("DROP TABLE {}".format(table))
        print("Table dropped... ")

        # Commit your changes in the database
        conn.commit()

        # Closing the connection
        conn.close()

    def delete_all_tasks(self, name=''):
        """
        Delete all rows in the tasks table
        :param conn: Connection to the SQLite database
        :return:
        """
        # Connecting to sqlite
        conn = sqlite3.connect('SQLite_Python.db')

        sql = 'DELETE FROM {}'.format(name)
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()

    def add_data(self, file='', name=''):
        conn = sqlite3.connect('SQLite_Python.db')
        c = conn.cursor()

        read_parts1 = pd.read_csv(r'{}'.format(file))
        read_parts1.to_sql('{}'.format(name), conn, if_exists='replace',
                            index=False)  # Replace the values from the csv file into the table 'name'

    def retrieve_data(self, data="", table="", row="", value=""):
        conn = sqlite3.connect('SQLite_Python.db')
        cursor = conn.cursor()
        cursor.execute("SELECT {} FROM {} WHERE {} = ?".format(data, table, row), (value,))
        rows = cursor.fetchall()

        data = [new for row in rows for new in row]
        return data

    def retrieve_all(self, data="", table=''):
        conn = sqlite3.connect('SQLite_Python.db')
        cursor = conn.cursor()
        cursor.execute("SELECT {} FROM {}".format(data, table))

        rows = cursor.fetchall()
        data = [new for row in rows for new in row]
        return data

    def retrieve_columns(self, data=None, table=''):
        conn = sqlite3.connect('SQLite_Python.db')
        cursor = conn.cursor()
        cursor.execute("SELECT {} FROM {}".format(data, table))

        rows = cursor.fetchall()
        data = [new for row in rows for new in row]
        return data

def main():
    db = Database()
    db.delete_table('CHITIET')
    db.create_parts1_table('CHITIET')
    db.add_data('example_db/file_thu.csv', 'CHITIET')
    all = db.retrieve_all('11524', 'CHITIET')
    soluong = db.retrieve_data('so_luong', 'CHITIET', "COM_11524", '1')
    print(soluong)
    # db.retrieve_all('*', 'Parts1')


if __name__ == '__main__':
    main()
