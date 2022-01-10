import sys
import os
print(os.path.abspath(''))
sys.path.append(f'{os.path.abspath("")}\\venv')
sys.path.append(f'{os.path.abspath("")}\\venv\\Scripts')
sys.path.append(f'{os.path.abspath("")}\\venv\\Lib\\site-packages')

from pickle import PROTO
import traceback
import pandas as pd
import win32com.client as win32


class Subiekt:
##########################################################################################
    # Uruchom(TypDopasowania, TypUruchomienia)
    # 
    #     TypDopasowania
    #         gtaUruchomDopasuj = 0
    #             Oznacza dopasowanie pierwszej znalezionej aplikacji zadanego typu,
    #             która jest podłączona do wskazanego serwera i podanej bazy danych.
    #         gtaUruchomDopasujUzytkownika = 1
    #             Oznacza dopasowanie pierwszej znalezionej aplikacji zadanego typu,
    #             która jest podłączona do wskazanego serwera z wykorzystaniem podanego
    #             użytkownika SQL Servera i podanej bazy danych.
    #         gtaUruchomDopasujOperatora = 2
    #             Oznacza dopasowanie pierwszej znalezionej aplikacji zadanego typu,
    #             która jest podłączona do wskazanego serwera,
    #             bazy danych oraz zalogowana na podanego użytkownika InsERT GT (operatora).

                
    #     TypUruchomienia
    #         gtaUruchom = 0
    #             Oznacza uruchomienie zadanej aplikacji o ile nie zostanie znaleziona działająca,
    #             do której można się podłączyć.
    #         gtaUruchomNieZablokowany = 1
    #             Oznacza uruchomienie zadanej aplikacji o ile nie zostanie znaleziona działająca
    #             aplikacja danego typu w stanie nie zablokowanym do której można się podłączyć.
    #             Tryb zablokowania ustawiany jest np. podczas uruchamiania aplikacji,
    #             kiedy inicjowane są połączenia, sprawdzane licencje itp.,
    #             także wówczas, gdy zmieniany jest użytkownik (Ctrl+F2).
    #             Proces programu Insert GT już istnieje w systemie,
    #             ale jeszcze nie jest gotowy do normalnej pracy.
    #         gtaUruchomNowy = 2 
    #             Oznacza zawsze uruchomienie nowej instancji aplikacji.
    #         gtaUruchomWTle = 4
    #             Oznacza uruchomienie/podłączenie do zadanej aplikacji
    #             bez wykorzystywania interfejsu użytkownika.
    #             Aplikacja podłączona w ten sposób działa w tle - nie otwiera się jej okno
#############################################################################################
    
    def __init__(self):
        self.oGT = win32.Dispatch('InsERT.GT')
        self.oGT.Produkt = 1
        self.oGT.Autentykacja = 0
        self.oGT.Serwer = "Server name"
        self.oGT.Uzytkownik = "user name"
        self.oGT.UzytkownikHaslo = "pasword name"
        self.oGT.Baza = "database name"
        self.oGT.Operator = "user login name for subiekt gt"
        self.oGT.OperatorHaslo = "user password for subiekt gt"

        self.error = []
    
    def __enter__(self):
        print('\n################### Program start    ########################')
        self.oSgt = self.oGT.Uruchom(2,2)
        if self.oSgt.Okno.Widoczne == False:
            self.oSgt.Okno.Widoczne = True
        
        return self
    
    def __exit__(self, exc_type, exc_value, exc_traceback):
        print('\n################### Program zakonczony ########################')
        self.oSgt.Zakoncz()
        print(f'ilośc błedów: {len(self.error)}')
        if len(self.error) > 0:
            for x in self.error:
                print(x)

    
    def zmiana_tw_gl_poboczna(self, czy_gl, numer_pob, symbol=''):

        try:
            if self.oSgt.Towary.Istnieje(symbol):
                self.oTw = self.oSgt.Towary.Wczytaj(symbol)
                print(self.oTw.Symbol)
                print(self.oTw.PoleWlasne('Główna kartoteka'))
                self.oTw.Zapisz()

            else:
                print('brak kartoteki')
                
        
        except Exception as e:
            print(e)


    def add_to_error(self, position, error):
        try:
            self.error.append((position, error))
        except Exception as e:
            print(e)

    def create_dok_zk(self):
        self.zk = self.oSgt.SuDokumentyManager.DodajPZ()
        for x in range(4):
            poz = self.zk.Pozycje.Dodaj(13179)
            poz.IloscJm = 10
            poz.Jm = '100szt.'
            poz.CenaNettoPrzedRabatem = 11.11
        
        self.zk.Wyswietl()

    def _tw_istnieje(self, towar):
        if self.oSgt.Towary.Istnieje(towar):
            return True
        else:
            self.add_to_error(towar, 'nie istnieje')
            return False
    
    def _tw_wczytaj(self, towar):
        try:
            if self._tw_istnieje(towar):
                self.oTw = self.oSgt.Towary.Wczytaj(towar)
                print(f'>>>>>>>>>>>>>>>>>>>>>WCZYTANY\nid : {self.oTw} - symbol: {self.oTw.Symbol}')
                print('----------------------')
                return True
            else:
                print('NIE_ISTNIEJE>>>>>>>>>>>>>>>>>')
                print(f'{towar} -> towar nie istnieje !!!!')
                print('NIE_ISTNIEJE<<<<<<<<<<<<<<<<<')
                return False
        except Exception as e:
            print('NIE_WCZYTANO>>>>>>>>>>>>>>>>>')
            print(f'{towar} -> towar nie WCZYTANO !!!!')
            print('NIE_WCZYTANO<<<<<<<<<<<<<<<<<')
            self.add_to_error(towar, e)

    def tw_zapisz(self):
        try:
            self.oTw.Zapisz()
            print('<<<<<<<<<<<<<<<<<<<<<ZAPISANY\n')
        except Exception as e:
            self.add_to_error('', e)

    def tw_zmien_symbol(self, towar, new_symbol, save=False):
        if self._tw_wczytaj(towar):
            try:
                if self.oTw.Symbol != new_symbol and len(new_symbol) <= 20:
                    print(f"old: {self.oTw.Symbol}\nnew: {new_symbol}")
                    self.oTw.Symbol = new_symbol
                    if save:
                        self.tw_zapisz()
                elif len(new_symbol) > 20:
                    self.add_to_error(towar, f'za dluga symbol - {len(new_symbol)}') 
            except Exception:
                print(f'symbol: {self.oTw.Symbol} -------------------Błąd')
                self.add_to_error(towar, 'bład tw_zmien_symbol') 


    def tw_zmien_nazwa(self, towar, new_nazwa, save=False):
        if self._tw_wczytaj(towar):
            try:
                if self.oTw.Nazwa != new_nazwa and len(new_nazwa) <= 50:
                    print(f"old: {self.oTw.Nazwa}\nnew: {new_nazwa}")
                    self.oTw.Nazwa = new_nazwa
                    if save:
                        self.tw_zapisz()       
                elif len(new_nazwa) > 51:
                    self.add_to_error(towar, f'za dluga nazwa - {len(new_nazwa)}') 
            except Exception:
                print(f'symbol: {self.oTw.Symbol} -------------------Błąd')
                self.add_to_error(towar, 'bład tw_zmien_nazwa') 
    
    def tw_zmien_opis(self, towar, new_opis, save=False, replace_opis=True):
        if self._tw_wczytaj(towar):
            try:
                if self.oTw.Opis != new_opis and len(new_opis) < 255:
                    if replace_opis is True:
                        self.oTw.Opis = new_opis
                        print(f"old: {self.oTw.Opis}\nnew: {new_opis}")
                    elif replace_opis is False:
                        self.oTw.Opis = f'{new_opis}\n {self.oTw.Opis}'
                        print(f"old: {self.oTw.Opis}\nnew: {new_opis}\n {self.oTw.Opis}")

                    if save:
                        self.tw_zapisz()       
                elif len(new_opis) >= 255:
                    self.add_to_error(towar, f'za dluga opis - {len(new_opis)}')
            except Exception:
                print(f'symbol: {self.oTw.Symbol} -------------------Błąd')
                self.add_to_error(towar, 'bład tw_zmien_opis') 

    def tw_zmien_pole2(self, towar, new_pole2, save=False):
        if self._tw_wczytaj(towar):
            try:
                if self.oTw.Pole2 != new_pole2 and len(new_pole2) < 255:
                    print(f"old: {self.oTw.Pole2}\nnew: {new_pole2}")
                    self.oTw.Pole2 = new_pole2
                    if save:
                        self.tw_zapisz()       
                elif len(new_pole2) >= 255:
                    self.add_to_error(towar, f'za dluga opis - {len(new_pole2)}') 
            except Exception:
                print(f'symbol: {self.oTw.Symbol} -------------------Błąd')
                self.add_to_error(towar, 'bład tw_zmien_pole2') 

    def dodaj_jm(self, towar, jm_name, jm_prze, save=False, second=False):
        if self._tw_wczytaj(towar):
            try:
                self.jm = self.oTw.Miary.Dodaj(jm_name) 
                self.jm.Przelicznik = jm_prze
                print(f"nazwa: {self.oTw.Symbol}\njm: {jm_name}\nprze: {jm_prze}")
                if second:
                    self.jm.JednostkaPorownawcza = True
                    print('Ustawiona jako jm_gł')
                if save:
                    self.tw_zapisz()
            except Exception:
                print(f'symbol: {self.oTw.Symbol} -------------------Błąd')
                self.add_to_error(towar, 'bład dodaj_jm') 

    def dodaj_WZv(self, kh_id: int, dok_title: str, dok_second_title: str, comments: str, df: pd.DataFrame):
        try:
            doc_wzv = self.oSgt.SuDokumentyManager.DodajWZv()
            doc_wzv.OdbiorcaId = kh_id
            doc_wzv.Tytul = dok_title
            doc_wzv.Podtytul = dok_second_title
            doc_wzv.Uwagi = comments

            for i, r in df.iterrows():
                # id describe quantity unit price
                poz = doc_wzv.Pozycje.Dodaj(int(r['id']))
                poz.Jm = r['unit']
                poz.IloscJm = r['quantity']
                poz.CenaNettoPrzedRabatem = r['price']
                poz.Opis = r['describe']

            # oDok.Zapisz()
            doc_wzv.Wyswietl()
            # oDok.Zamknij()

        except Exception:
            self.add_to_error(kh_id, 'bład dodaj_WZv') 
            traceback.print_exc()
