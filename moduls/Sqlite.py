import sys
import os
sys.path.append(f'{os.path.abspath("")}\\venv')
sys.path.append(f'{os.path.abspath("")}\\venv\\Scripts')
sys.path.append(f'{os.path.abspath("")}\\venv\\Lib\\site-packages')
import sqlite3
from sqlite3 import Error
from PyQt6.sip import transferback
import pandas as pd
import traceback


class Sqlite:

    def __init__(self, sqlite3_db=None) -> None:
        if sqlite3_db is None:
            sqlite3_db = 'pythonsqlite.db'
        self.sqlite3_db = sqlite3_db

    def __enter__(self):
        return self.create_connection(self.sqlite3_db)

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.close_connection()

    def create_connection(self, sqlite3_db):
        try:
            self.conn = sqlite3.connect(sqlite3_db)
            self.cur = self.conn.cursor()
            print(f'{sqlite3_db} - connect')
            return self
        except Error as e:
            traceback.print_exc()
            return None
    
    def close_connection(self):
        try:
            self.cur.close()
            print(f'{self.sqlite3_db} cursor close')
        except Error as e:
            traceback.print_exc()
        try:
            self.conn.close()
            print(f'{self.sqlite3_db} connection close')
        except Error as e:
            traceback.print_exc()

    def create_table(self, query):
        """ create a table from the create_table_sql statement
        :param conn: Connection object
        :param create_table_sql: a CREATE TABLE statement
        :return:
        """
        try:
            self.cur.execute(query)
        except Error as e:
            traceback.print_exc()

    def insert_record(self, an_DokId, an_DokNazwa, an_KhSym, an_Stany, an_Ceny, an_Czas):
        """Insert proceded dokument"""

        query = ''' INSERT INTO
                        AnalizaDokomentow(an_DokId,an_DokNazwa,an_KhSym,an_Stany,an_Ceny,an_Czas)
                        VALUES(?,?,?,?,?,?) '''
        
        values = [an_DokId, an_DokNazwa, an_KhSym, an_Stany, an_Ceny, an_Czas]
        try:
            self.cur.execute(query, values)
            self.conn.commit()
        except Exception as e:
            traceback.print_exc()

    def select_all(self):
        query = "SELECT * FROM AnalizaDokomentow"
        columns = ['an_Id','an_DokId','an_DokNazwa','an_KhSym','an_Stany','an_Ceny','an_Czas']
        try:
            self.cur.execute(query)
            df = self._to_df(self.cur.fetchall(), columns)
            df['an_Czas'] = pd.to_datetime(df['an_Czas'])
            return df
        except Error as e:
            traceback.print_exc()

    def select_all_doc_ids(self) -> list:
        query = "SELECT an_DokId FROM AnalizaDokomentow"
        try:
            self.cur.execute(query)
            an_DokId = [x[0] for x in self.cur.fetchall()]
            return an_DokId
        except Error as e:
            traceback.print_exc()

    def _to_df(self, list_of_dict: list, columns: list=None) -> pd.DataFrame:
        try:
            df = pd.DataFrame(list_of_dict)
            if columns is not None and df.shape[1] == len(columns):
                df.columns = columns
            return df
        except Exception as e:
            traceback.print_exc()      
