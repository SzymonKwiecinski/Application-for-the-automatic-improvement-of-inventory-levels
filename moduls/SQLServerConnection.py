import sys
import os
sys.path.append(f'{os.path.abspath("")}\\venv')
sys.path.append(f'{os.path.abspath("")}\\venv\\Scripts')
sys.path.append(f'{os.path.abspath("")}\\venv\\Lib\\site-packages')

import pyodbc
import pandas as pd
import re

class SQLServerConnection:
    """Connects with MS SQL Server and manage it.
    
    Attributes:
        conn (pyodbc.connect)
        cursor (pydoc.connect.coursor)

    """

    def __init__(self, db_log=None) -> None:
        if db_log is None:
            self.db_log = { 'Driver': r'{SQL Server Native Client 11.0}',
                            'Server': r'database server ip',
                            'Database': r'databes name',
                            'Uid': r'user name',
                            'PWD': r'pasword name'}
        else:
            self.db_log = db_log

    def __enter__(self) -> object:
        self.conn =\
                pyodbc.connect(f"Driver={self.db_log['Driver']};"
                               f"Server={self.db_log['Server']};"
                               f"Database={self.db_log['Database']};"
                               f"Uid={self.db_log['Uid']};"
                               f"PWD={self.db_log['PWD']};")
        self.cursor = self.conn.cursor()
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.cursor.close()
        self.conn.close()

    def to_df(self, columns=None, sql_str=None) -> pd.DataFrame:
        try:
            self.query(sql_str)
            return self._list_of_dicts_to_dataframe(columns)
        except:
            raise os.error


    def query(self, sql_str):
        try:
            self.cursor.execute(sql_str)
        except Exception as e:
            print(f'query fail \t {e}\n')

    def insert(self, sql_str):
        if 'INSERT' in sql_str.upper():
            try:
                self.query(sql_str=sql_str)
                self.conn.commit()
                return True
            except Exception:
                return False
    
    def delete(self, sql_str):
        return self.insert(sql_str)

    def _query_to_list(self):
        try:
            return [item[0] for item in self.cursor]
        except Exception as e:
            print(f'_query_to_list fail \t {e}\n')
            return None


    def _cursor_to_list_of_dicts(self) -> list:
        try:
            list_of_dict = []
            for row in self.cursor:
                d_poz = { i_x:row_x  for i_x, row_x in enumerate(row)}
                list_of_dict.append(d_poz)   
            return list_of_dict
        except Exception as e:
            print(f'_cursor_to_list_of_dicts fail \t {e}\n')
            return None

    def _list_of_dicts_to_dataframe(self, columns=None) -> pd.DataFrame:
        try:
            df = pd.DataFrame(self._cursor_to_list_of_dicts())
            if columns is not None and len(df.columns) == len(columns):
                df.columns = columns
            return df
        except Exception as e:
            print(f'_list_of_dicts_to_dataframe fail \t {e}\n')
            return None

    def _sql_to_str(self, path_file: str) -> str:
        with open(path_file) as file:
            return ''.join(file.readlines())
    