import sys
import os
sys.path.append(f'{os.path.abspath("")}\\venv')
sys.path.append(f'{os.path.abspath("")}\\venv\\Scripts')
sys.path.append(f'{os.path.abspath("")}\\venv\\Lib\\site-packages')

import traceback
from moduls.SQLServerConnection import SQLServerConnection
from moduls.CheckUnits import CheckUnit
import pandas as pd
from datetime import date, datetime, timedelta
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.QtCore import *
from moduls.Sqlite import Sqlite
from moduls.Subiekt import Subiekt

class Query:
    def __init__(self) -> None:
        pass

class Positions(Query):
    """Class manage DataFrame with positions to correct.
    
    Args:
        query (str): SQL query if None then take all exist docs
    
    Attributes:
        query (str): SQL query
            dok_id, dok_nr, dok_date, lp, pod_id, pod_sym, pod_naz,
            ilosc, jm, cena_netto, czy_gl, gl_id, gl_sym, gl_naz,
            kh_id, kh_sym, dok_status
        df (DataFrame): results of query
        columns (list): name of columns in DataFrame
        unify_df (DataFrame): unified DataFrame 
        wrong_list (list): list of units without pair in units 
            with main positions in system
    """
    def __init__(self, query: str=None) -> None:
        super().__init__()
        if query is None:
            self.query = """
                    SELECT
                        d.dok_Id as dok_id,
                        d.dok_NrPelny as dok_nr,
                        d.dok_DataWyst as dok_date,
                        p.ob_DokHanLp as lp,
                        t.tw_Id as pob_id,
                        t.tw_Symbol as pob_sym,
                        t.tw_Nazwa as pob_naz,
                        p.ob_Ilosc as ilosc,
                        p.ob_Jm as jm,
                        p.ob_CenaNetto as cena_netto,
                        pd.pwd_Flaga01 as czy_gl,
                        pd.pwd_Liczba01 as gl_id,
                        (SELECT tw_Symbol FROM AREX.dbo.tw__Towar WHERE tw_Id = pd.pwd_Liczba01) as gl_sym,
                        (SELECT tw_Nazwa FROM AREX.dbo.tw__Towar WHERE tw_Id = pd.pwd_Liczba01) as gl_naz,
                        d.dok_PlatnikId as kh_id,
                        k.kh_Symbol as kh_sym,
                        d.dok_Status as dok_status

                    FROM AREX.dbo.dok_Pozycja as p
                        LEFT JOIN AREX.dbo.dok__Dokument as d
                            ON p.ob_DokHanId = d.dok_Id
                        LEFT JOIN AREX.dbo.tw__Towar as t
                            ON p.ob_TowId = t.tw_Id
                        LEFT JOIN AREX.dbo.kh__Kontrahent as k
                            ON d.dok_PlatnikId = k.kh_Id
                        LEFT JOIN AREX.dbo.pw_Dane as pd
                            ON t.tw_Id = pd.pwd_IdObiektu
                            AND pd.pwd_TypObiektu = -14

                    WHERE d.dok_Typ = 2
                    AND pd.pwd_Flaga01 = 0
                    AND k.kh_Id IN (78,339,535,3482,508,795)
                    ORDER BY d.dok_Id DESC
                            """
        if query is not None:
            self.query = query

        self.columns =   ['dok_id', 'dok_nr', 'dok_date', 'lp', 'pod_id',
                    'pod_sym', 'pod_naz', 'ilosc', 'jm','cena_netto',
                    'czy_gl', 'gl_id', 'gl_sym', 'gl_naz', 'kh_id',
                    'kh_sym','dok_status']

        self.df = self.all_positions
        self.exclude_positions()
        self.set_data_types()
        self.units_from_1pieces_to_100pieces()

    @property
    def all_positions(self) -> pd.DataFrame:
        with SQLServerConnection() as sql:
            df = sql.to_df(sql_str=self.query, columns=self.columns)
        return df

    def exclude_positions(self) -> None:
        """Exclude done before records from DataFrame."""
        with Sqlite() as sql:
            exc_doc_ids = sql.select_all_doc_ids()
        if len(self.df) > 0:
            try:
                self.df = self.df[~self.df['dok_id'].isin(exc_doc_ids)]
            except:
                traceback.print_exc()

    @classmethod
    def from_date(clc, date_start: str, date_end: str):
        """Create df filter by date.
        
        Args:
            date_start (str)
            date_end (str)
        
        Returns:
            class Positions
        """
        query = f"""
            SELECT
                d.dok_Id as dok_id,
                d.dok_NrPelny as dok_nr,
                d.dok_DataWyst as dok_date,
                p.ob_DokHanLp as lp,
                t.tw_Id as pob_id,
                t.tw_Symbol as pob_sym,
                t.tw_Nazwa as pob_naz,
                p.ob_Ilosc as ilosc,
                p.ob_Jm as jm,
                p.ob_CenaNetto as cena_netto,
                pd.pwd_Flaga01 as czy_gl,
                pd.pwd_Liczba01 as gl_id,
                (SELECT tw_Symbol FROM AREX.dbo.tw__Towar WHERE tw_Id = pd.pwd_Liczba01) as gl_sym,
                (SELECT tw_Nazwa FROM AREX.dbo.tw__Towar WHERE tw_Id = pd.pwd_Liczba01) as gl_naz,
                d.dok_PlatnikId as kh_id,
                k.kh_Symbol as kh_sym,
                d.dok_Status as dok_status

            FROM AREX.dbo.dok_Pozycja as p
                LEFT JOIN AREX.dbo.dok__Dokument as d
                    ON p.ob_DokHanId = d.dok_Id
                LEFT JOIN AREX.dbo.tw__Towar as t
                    ON p.ob_TowId = t.tw_Id
                LEFT JOIN AREX.dbo.kh__Kontrahent as k
                    ON d.dok_PlatnikId = k.kh_Id
                LEFT JOIN AREX.dbo.pw_Dane as pd
                    ON t.tw_Id = pd.pwd_IdObiektu
                    AND pd.pwd_TypObiektu = -14

            WHERE d.dok_Typ = 2
            AND d.dok_DataWyst >= '{date_start}'
            AND d.dok_DataWyst <= '{date_end}'
            AND pd.pwd_Flaga01 = 0
            AND k.kh_Id IN (78,339,535,3482,508,795)
            ORDER BY d.dok_Id DESC
                    """
        
        return clc(query)
        
    @classmethod
    def from_dok_id(clc, dok_id: int):
        """Create df filter by document id.
        
        Args:
            dok_id (int)
        
        Returns:
            class Positions
        """
        query = f"""
            SELECT
                d.dok_Id as dok_id,
                d.dok_NrPelny as dok_nr,
                d.dok_DataWyst as dok_date,
                p.ob_DokHanLp as lp,
                t.tw_Id as pob_id,
                t.tw_Symbol as pob_sym,
                t.tw_Nazwa as pob_naz,
                p.ob_Ilosc as ilosc,
                p.ob_Jm as jm,
                p.ob_CenaNetto as cena_netto,
                pd.pwd_Flaga01 as czy_gl,
                pd.pwd_Liczba01 as gl_id,
                (SELECT tw_Symbol FROM AREX.dbo.tw__Towar WHERE tw_Id = pd.pwd_Liczba01) as gl_sym,
                (SELECT tw_Nazwa FROM AREX.dbo.tw__Towar WHERE tw_Id = pd.pwd_Liczba01) as gl_naz,
                d.dok_PlatnikId as kh_id,
                k.kh_Symbol as kh_sym,
                d.dok_Status as dok_status

            FROM AREX.dbo.dok_Pozycja as p
                LEFT JOIN AREX.dbo.dok__Dokument as d
                    ON p.ob_DokHanId = d.dok_Id
                LEFT JOIN AREX.dbo.tw__Towar as t
                    ON p.ob_TowId = t.tw_Id
                LEFT JOIN AREX.dbo.kh__Kontrahent as k
                    ON d.dok_PlatnikId = k.kh_Id
                LEFT JOIN AREX.dbo.pw_Dane as pd
                    ON t.tw_Id = pd.pwd_IdObiektu
                    AND pd.pwd_TypObiektu = -14

            WHERE d.dok_Typ = 2
            AND d.dok_Id >= {dok_id}
            AND pd.pwd_Flaga01 = 0
            AND k.kh_Id IN (78,339,535,3482,508,795)
            ORDER BY d.dok_Id DESC
            """
        
        return clc(query)

    def check_status(self) -> list:
        """Check document by magazine realise status.
        
        If 0 bad, if 1 good.
        
        Returns:
            list: list of bad document or empty list if
                everthing is good
        """
        wrong_docs = [x for x in self.df[self.df['dok_status'] == 0]['dok_nr'].unique()]
        return wrong_docs

    def set_data_types(self)  -> None:
        """Sets date types of DataFrame."""
        if len(self.df) > 0 :
            self.df = self.df.astype(
                {
                    'dok_id': 'int64',
                    'dok_nr': 'object',
                    'dok_date': 'datetime64',
                    'lp': 'int64',
                    'pod_id': 'int64',
                    'pod_sym': 'object',
                    'pod_naz': 'object',
                    'ilosc': 'float64',
                    'cena_netto': 'float64',
                    'czy_gl': 'object',
                    'gl_id': 'int64',
                    'gl_sym': 'object',
                    'gl_naz': 'object',
                    'kh_id': 'int64',
                    'kh_sym': 'object',
                    'dok_status': 'int64'
                }
            )

    def units_from_1pieces_to_100pieces(self) -> None:
        """Chenge units in DataFrame from 'szt.' to '100szt.'"""

        for i ,r in self.df.iterrows():
            if r['jm'] == 'szt.':
                self.df.loc[i, 'jm'] = '100szt.'
                self.df.loc[i, 'ilosc'] = r['ilosc'] / 100
                self.df.loc[i, 'cena_netto'] = r['cena_netto'] / 100

    def unify(self, df_units: pd.DataFrame) -> bool:
        """Check if all measure if units are in system.
        
        Create new column 'unify'=False. If find
        pair then chenge df['unify'] = True

        Args:
            df_units (str): DataFrame with units from system
        
        Returns:
            bool:
                True: if all positions have pair of units in system
                False: if at least one do not have pair in system
        """
        self.unify_df = self.df

        self.unify_df['unify'] = False
        for i, r in self.unify_df.iterrows():
            merge_id = df_units[df_units['id'] == r['gl_id']]
            if len(merge_id) > 0:
                merge_unit = merge_id[merge_id['jm'] == r['jm']]
                if len(merge_unit) == 1:
                    self.unify_df.loc[i,'unify'] = True
        
        if len(self.unify_df[self.unify_df['unify'] == False]) == 0:
            return True
        else:
            self.get_wrong_postitions()
            return False

    def get_wrong_postitions(self) -> None:
        """Retrive positions without pair in system by units."""
        wrong_df = self.unify_df[self.unify_df['unify'] == False]
        self.wrong_list = []
        for i, r in wrong_df.iterrows():
            self.wrong_list.append({
                'symbol': r['gl_sym'],
                'jednostka_miary': r['jm']
                })

    def group_by_contrahent(self) ->list:
        """Return DataFrame split by contrahent in list od disc.
        
        Returns:
            list: list of disct ({'kh_sym': str,'df': DataFrame})"""

        contrahents_sym = self.unify_df['kh_sym'].unique()
        contrahents_df = []
        for contrahent in contrahents_sym:
            contrahents_df.append({
                'kh_sym': contrahent,
                'df': self.unify_df[self.unify_df['kh_sym'] == contrahent]
                })
        return contrahents_df

class Units(Query):
    """Class manage DataFrame with all units from system.
    
    Precisely all main and auxiliary units from main
    positions in system  

    Args:
        query (str): SQL query if None then take all exist docs
    
    Attributes:
        query (str): SQL query
            id, symbol, nazwa, jm, przelicznik, czy_gl_jm
        df (DataFrame): results of query
        columns (list): name of columns in DataFrame
    """
    def __init__(self) -> None:
        super().__init__()

        self.columns = ['id', 'symbol', 'nazwa', 'jm', 'przelicznik', 'czy_gl_jm']
        self.query = """WITH jm_gl as(
            SELECT
                t.tw_Id as id,
                t.tw_Symbol as symbol,
                t.tw_Nazwa as nazwa,
                t.tw_JednMiary as jm,
                1.0 as przelicznik,
                1 as czy_gl_jm

            FROM AREX.dbo.tw__Towar AS t
            LEFT JOIN AREX.dbo.pw_Dane as pd
                ON t.tw_Id = pd.pwd_IdObiektu
                AND pd.pwd_TypObiektu = -14
            WHERE
                t.tw_Zablokowany = 0
                AND pd.pwd_Flaga01 = 1
            ),


            jm_pob as(
                SELECT
                    t.tw_Id as id,
                    t.tw_Symbol as symbol,
                    t.tw_Nazwa as nazwa,
                    jm.jm_IdJednMiary as jm,
                    jm.jm_Przelicznik as przelicznik,
                    0 as czy_gl_jm

                FROM AREX.dbo.tw__Towar AS t
                FULL JOIN AREX.dbo.tw_JednMiary AS jm
                    ON t.tw_Id = jm.jm_IdTowar
                LEFT JOIN AREX.dbo.pw_Dane as pd
                    ON t.tw_Id = pd.pwd_IdObiektu
                    AND pd.pwd_TypObiektu = -14
                WHERE
                    t.tw_Zablokowany = 0
                    AND jm.jm_IdTowar is not null
                    AND pd.pwd_Flaga01 = 1
            )

            SELECT * 
            FROM jm_gl
            UNION
                SELECT *
                FROM jm_pob 
            ORDER BY symbol DESC"""
        self.df = self.all_units

    @property
    def all_units(self) -> pd.DataFrame:
        with SQLServerConnection() as sql:
            df = sql.to_df(sql_str=self.query, columns=self.columns)
        return df

class Document():
    """Class create doc from DataFrame and export to Subiekt.
    
    Create document WZv based of DataFrame and using 'Sfera',
    create document in subiekt

    Args:
        df (DataFrame): DataFrame with data for one company

    Attributes:
        df (DataFrame): DataFrame which we work on
        document_df (DataFrame): DataFrame fits to export do Subiekt
            We work main on it
        all_doks_id (list): list of all doc_id in document 
        doc_name_for_max_id (str): Doc full name for highest id number id DataFrame
        kh_id (int): numer id in Subiekt for the contrahent
        kh_sym (str): symbol in Subiekt for the contrahent
        date_range (dict): {'start': datetime ,'end': datetime}
            Date range in DataFrame for document 
        dok_title (str): 'Wydanie zewnętrzne z VAT Automatyczne'
        dok_second_title (str): 'Wyrównianie stanów magazynowych'
        comments (str): string representation of date_range



    """

    def __init__(self, df: pd.DataFrame) -> None:
        super().__init__()
        self.df = df.reset_index(drop=True)
        self.set_dok_info()
        self.create_document()

    @property
    def all_doks_id(self) -> list:
        """Returns all doks_id which were base for new document."""
        doks = [x for x in self.df['dok_id'].unique()]
        return doks

    @property
    def doc_name_for_max_id(self) -> str:
        doc_name = set(self.df[self.df['dok_id'] == max(self.all_doks_id)]['dok_nr']).pop()
        return doc_name

    def dok_nr_for_dok_id(self, dok_id: int) -> str:
        doc_name = set(self.df[self.df['dok_id'] == dok_id]['dok_nr']).pop()
        return doc_name

    def data_for_dok_id(self, dok_id: int) -> datetime:
        doc_date = set(self.df[self.df['dok_id'] == dok_id]['dok_date']).pop()
        return doc_date

    def set_dok_info(self) -> None:
        """Sets basic info to document.
        
        Sets contrahent, title, commetns to new document
        """
        try:
            self.kh_id = int(self.df.loc[0,'kh_id'])
            self.kh_sym = str(self.df.loc[0,'kh_sym'])
            self.date_range = {'start': min(self.df['dok_date']) ,'end': max(self.df['dok_date'])}
            self.dok_title = 'Wydanie zewnętrzne z VAT Automatyczne'
            self.dok_second_title = 'Wyrównianie stanów magazynowych'
            self.comments = f'''Dni Od {self.date_range["start"]} Do {self.date_range["end"]}'''
        except Exception as e:
            print(e)

    def append_positions_on_doc(self) -> list:
        """Takes data from df and create positions on new document_df.
        
        Returns:
            list of dict
        """
        positions_on_doc = []
        try:
            for i ,r in self.df.iterrows():
                positions_on_doc.append(
                    {
                        'id': r['gl_id'],
                        'describe': f"{r['dok_nr']} \npos:{r['lp']} \nstary sym: {str(r['pod_sym'])} \n{str(r['pod_naz'])}",
                        'quantity': r['ilosc'],
                        'unit': r['jm'],
                        'price': r['cena_netto']
                    })
            return positions_on_doc
        except Exception:
            traceback.print_exc()

    def create_document(self) -> None:
        """Create document_df (df fits to export do Subiekt)."""
        positions_on_doc = self.append_positions_on_doc()
        self.document_df = pd.DataFrame(
            positions_on_doc,
            columns=['id','describe', 'quantity', 'unit', 'price'])
        self.document_df = self.document_df.astype(
            {
                'id': 'int64',
                'describe': 'object', 
                'quantity': 'float64', 
                'unit': 'object', 
                'price': 'float64'
            })

    def to_subiekt(self) -> None:
        """Export Dataframe to Subiekt.
        
        Create document in Subiekt based on
        document_df
        """

        with Subiekt() as sub:
            sub.dodaj_WZv(
                kh_id=self.kh_id,
                dok_title=self.dok_title,
                dok_second_title=self.dok_second_title,
                comments=self.comments, 
                df=self.document_df)

    def __str__(self) -> str:
        string =\
            f"""Inforamcje na tem dokumentu
            kh_id: {type(self.kh_id)} {self.kh_id}
            kh_sym: {type(self.kh_sym)} {self.kh_sym}
            dok_title: {type(self.dok_title)} {self.dok_title}
            dok_second_title: {type(self.dok_second_title)} {self.dok_second_title}
            comments: {type(self.comments)} {self.comments}
            
            Pozycje:
            {self.document_df.head()}"""

        return string

class Window(QWidget):
    """Main grafical class inherits from PyQt.QWidgets."""

    def __init__(self) -> None:
        super().__init__()
        self.ui()
        self.set_up()
        self.show()

    def ui(self):
        """Creates grafical component od program."""

        self.setWindowTitle('Stany magazynowe')
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.main_frame = QFrame()
        self.layout.addWidget(self.main_frame)
        self.layout_main_frame = QVBoxLayout()
        self.main_frame.setLayout(self.layout_main_frame)
        
        self.frame_1 = QFrame()
        self.layout_frame_1 = QFormLayout()
        self.frame_1.setLayout(self.layout_frame_1)
        self.layout_main_frame.addWidget(self.frame_1)
        self.layout_frame_1.addRow(QLabel('Ostatnio poprawiona faktura'))
        self.label_dok_f1 = QLabel('aa')
        self.label_dok_time_f1 = QLabel('aa')
        self.layout_frame_1.addRow(QLabel('Dokument: '), self.label_dok_f1)
        self.layout_frame_1.addRow(QLabel('Data wystawienia: '), self.label_dok_time_f1)

        self.frame_2 =QFrame()
        self.layout_frame_2 = QFormLayout()
        self.frame_2.setLayout(self.layout_frame_2)
        self.layout_main_frame.addWidget(self.frame_2)
        self.layout_frame_2.addRow(QLabel('Wybierz dni do poprawy stanów'))
        self.time_start_f2 = QDateEdit()
        self.time_end_f2 = QDateEdit()
        self.time_end_f2.setDate(datetime.today())
        self.time_end_f2.setMaximumDate(datetime.today())
        self.layout_frame_2.addRow(QLabel('Od: '), self.time_start_f2)#, QLabel(' do: '), self.time_end_f2) 
        self.layout_frame_2.addRow(QLabel('Do: '), self.time_end_f2)

        self.frame_3 =QFrame()
        self.layout_frame_3 = QVBoxLayout()
        self.frame_3.setLayout(self.layout_frame_3)
        self.layout_main_frame.addWidget(self.frame_3)
        self.list_w_f3 = QListWidget()
        self.btn_1_f3 = QPushButton('Popraw')
        self.btn_2_f3 = QPushButton('Drukuj raport')
        self.pro_bar_f3 = QProgressBar()
        self.pro_bar_f3.setMaximum(100)
        self.layout_frame_3.addWidget(QLabel('Informacje o konwersji'))
        self.layout_frame_3.addWidget(self.list_w_f3)
        self.layout_frame_3.addWidget(self.btn_1_f3)
        self.layout_frame_3.addWidget(self.pro_bar_f3)
        self.layout_frame_3.addWidget(self.btn_2_f3)

        self.btn_1_f3.clicked.connect(self.run_conversion)

    def add_item(self, item, status: bool=True):
        """Add info to list_widget on top of stack.
        
        When we send correct info then status=True gives green color
        When we send incorrect info then status=False gives red color
        
        Args:
            item (str): text of message
            status (bool): color of message
        """
        if status is True:
            color = '#82ed9d'
        elif status is False:
            color = '#c93a4c'
        else:
            raise TypeError
        item = QListWidgetItem(item)
        item.setBackground(QColor(color))
        self.list_w_f3.addItem(item)

    def set_up(self):
        """Set up propgram variable when program turn on."""
        self.units = Units()
        self.find_last_record()

    def find_last_record(self):
        """Finds last converted ID of document.
        
        Connect with SQLITE database and retrives info about
        last dokument order by Dok_Id and set these info to GUI.
        """
        with Sqlite() as sql:
            df = sql.select_all()
        try:
            df = df.nlargest(columns=['an_DokId'], n=1).reset_index(drop=True)
            dok_name = df.loc[0,'an_DokNazwa']
            dok_time = df.loc[0,'an_Czas']
            self.label_dok_f1.setText(dok_name)
            self.label_dok_time_f1.setText(dok_time.strftime('%Y-%m-%d'))
            self.time_start_f2.setDate(dok_time)
        except Exception as e:
            print(e)

    def results_to_sqlite(self, answer: bool, document):
        """Saves docs id number to sqlite db.
        
        Args:
            answer (bool): if True doc was save
                if False doc was not save!
        """
        
        if answer is True:
            try:
                self.add_item(f'{document.kh_sym} - WZv wykonane pomyślnie', True)
                if len(document.all_doks_id) != 0:
                    with Sqlite() as sql:
                        for x in document.all_doks_id:
                            sql.insert_record(
                                    an_DokId=int(x),
                                    an_DokNazwa=document.dok_nr_for_dok_id(int(x)),
                                    an_KhSym=document.kh_sym,
                                    an_Stany='Tak',
                                    an_Ceny='Nie',
                                    an_Czas=str(document.data_for_dok_id(int(x)))
                                )
            except Exception:
                traceback.print_exc()
        elif answer is False:
            self.add_item(f'Błąd - {document.kh_sym} - WZv nie zrobione', False)



        # doc.to_subiekt(self)

    def check_date(self) -> bool:
        """Check Does start date is not later then end date."""

        if self.time_start_f2.date() <= self.time_end_f2.date():
            self.add_item('Czas się zgadza', True)
            return True
        else:
            self.add_item('Bład - zły czas!', False)
            return False

    def run_conversion(self):
        """Runs sequence of methods to convert positions/goods.
        
        1.Retrive auxiliary positions filter by date
        2.Chceck if all docs have correct release status
        3.Unify auxiliary positions to main positions
        4.Create Documents
        5.Export Doxuments to Subiekt
        """
        if self.check_date() is True:
            self.list_w_f3.clear()
            self.positions = Positions.from_date(self.time_start_f2.text(), self.time_end_f2.text())
            
            if len(self.positions.df) > 0:
                self.wrong_docs = self.positions.check_status()
                if len(self.wrong_docs) > 0:
                    QMessageBox.information(self,'Info',f'Proszę zrealizować następujące dokumenty\n{self.wrong_docs}')
                    for x in self.wrong_docs:
                        self.add_item(f'Błąd - zrealizuj wydanie magazynowe dla {x}', False)
                else:
                    self.add_item('Wszystkie dokumenty są prawidłowe', True)
                    self.pro_bar_f3.setValue(25)

                    status = self.positions.unify(self.units.df)
                    if status is False:
                        QMessageBox.information(self,'Info',f'Poprawić poboczne jednostki miary\n{self.positions.wrong_list}')
                        for x in self.positions.wrong_list:
                            self.add_item(f'Błąd - Dodaj do {x["symbol"]} jednostkę {x["jm"]}', False)
                    elif status is True:
                        self.add_item('Wszystkie kartoteki są poprawne', True)
                        self.pro_bar_f3.setValue(50)
                        df_split_by_kh = self.positions.group_by_contrahent()

                        if len(df_split_by_kh) > 0:
                            counter_to_bar = int(50 / len(df_split_by_kh))


                            for dictonary in df_split_by_kh:
                                dok  = Document(dictonary['df'])
                                dok.to_subiekt()
                                self.pro_bar_f3.setValue(self.pro_bar_f3.value() + counter_to_bar)
                                answer = QMessageBox.question(self ,'Zapytanie',f'Czy udało się zapisać dokument?\n{dok.kh_sym}',buttons=QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)
                                if answer == QMessageBox.StandardButton.Yes:
                                    self.results_to_sqlite(True, dok)
                                elif answer == QMessageBox.StandardButton.No or answer == QMessageBox.StandardButton.Close:
                                    self.results_to_sqlite(False, dok)

                                self.set_up()
                                self.pro_bar_f3.setValue(100)

                        elif len(df_split_by_kh) == 0:
                            self.add_item('Nie znaleziono kartotek do konwersji', True)
                            self.pro_bar_f3.setValue(100)
            else:
                self.add_item('Nie znaleziono kartotek do konwersji', True)
                self.pro_bar_f3.setValue(100)

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = Window()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()