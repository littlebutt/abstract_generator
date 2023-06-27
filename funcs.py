from globals import Global


class Func:

    def run(self, *args, **kwargs):
        pass

    @staticmethod
    def get_row_num(num: str) -> int:
        res = int(num) - 1
        if res < 0:
            raise RuntimeError("Bad parameter in 'SUM' function")
        return res

    @staticmethod
    def get_col_num(num: str) -> int:
        res = 0
        for c in num:
            res *= 26
            res += ord(c) - ord('A') + 1
        return res - 1


class SumFunc(Func):

    def __init__(self, excel_scanner: "ExcelScanner"):
        self.excel_scanner = excel_scanner

    def run(self, *args, **kwargs):
        if 'sheets' not in kwargs or 'col' not in kwargs or 'row' not in kwargs:
            raise RuntimeError("Missing args when calling 'SUM' function")
        res = 0.0
        for sheet in kwargs['sheets']:
            _data = self.excel_scanner.read(self.get_row_num(kwargs['row']), self.get_col_num(kwargs['col']), sheet)
            res += float(_data)
        if res.is_integer():
            return str(int(res))
        else:
            return '%.2f' % res


class NormFunc(Func):

    def __init__(self, excel_scanner: "ExcelScanner"):
        self.excel_scanner = excel_scanner

    def run(self, *args, **kwargs):
        if 'col' not in kwargs or 'row' not in kwargs:
            raise RuntimeError("Missing args when calling 'SUM' function")
        sheet = kwargs['sheet'] if 'sheet' in kwargs else Global.sheets[-1]
        res = self.excel_scanner.read(self.get_row_num(kwargs['row']), self.get_col_num(kwargs['col']), sheet)
        if res.is_integer():
            return str(int(res))
        else:
            return '%.2f' % res
