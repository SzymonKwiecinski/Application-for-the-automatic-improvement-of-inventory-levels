UNITS = {'SZT':     ['SZT', 'SZT.', '.SZT'],
         '10SZT':   ['10SZT', '10SZT.', '10.SZT', '10 SZT',
                     'SZT10', 'SZT.10', 'SZT10.', 'SZT 10'],
         '100SZT':  ['100SZT', '100SZT.', '100.SZT', '100 SZT',
                     'SZT100', 'SZT.100', 'SZT100.', 'SZT 100'],
         'KG':      ['KG', 'KG.'],
         'MB':      ['MB', 'MB.']}


class CheckUnit:
    """Unify units !!. Case no sensitive.
    
    Attributes:
        checking_unit (str): unit of measure which we want to unify

    """

    def __init__(self, checking_unit: str) -> str:
        self.checking_unit = checking_unit.upper()

    def unify(self):
        """Unify unit.
        
        Returns:
            str: unified unit with a capital letter

        """

        if self.checking_unit in UNITS['SZT']:
            unify_checking_unit = 'SZT'
        elif self.checking_unit in UNITS['10SZT']:
            unify_checking_unit = '10SZT'
        elif self.checking_unit in UNITS['100SZT']:
            unify_checking_unit = '100SZT'
        elif self.checking_unit in UNITS['KG']:
            unify_checking_unit = 'KG'
        elif self.checking_unit in UNITS['MB']:
            unify_checking_unit = 'MB'
        elif self.checking_unit == 'BRAK':
            unify_checking_unit = 'BRAK'
        else:
            raise Unify_unit_error('Błąd unifikacji jednostek')

        return unify_checking_unit.lower() + '.'

class Unify_unit_error(Exception):
    pass