import re

from globals import Global


class Func:

    def run(self, target: str) -> str:
        pass

    @classmethod
    def match(cls, target: str) -> bool:
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

    def run(self, target: str) -> str:
        target = target.strip()
        target = target[3:]
        target = target[1:-1]
        row, col = target.split(',')
        res = 0.0
        for sheet in Global.sheets:
            _data = self.excel_scanner.read(self.get_row_num(row.strip()), self.get_col_num(col.strip()), sheet)
            res += float(_data)
        if res.is_integer():
            return str(int(res))
        else:
            return '%.2f' % res

    @classmethod
    def match(cls, target: str) -> bool:
        return re.match(r'SUM\(.*\)', target.strip()) is not None


class NormFunc(Func):

    def __init__(self, excel_scanner: "ExcelScanner"):
        self.excel_scanner = excel_scanner

    def run(self, target: str) -> str:
        target = target.strip()
        target_list = target.split(',')
        if len(target_list) == 2:
            row, col = target_list
            sheet = Global.sheets[-1]
        elif len(target_list) == 3:
            row, col, sheet = target_list
        else:
            raise RuntimeError("Poor Argument count in NormFunc")
        res = self.excel_scanner.read(self.get_row_num(row.strip()), self.get_col_num(col.strip()), sheet.strip())
        if res.is_integer():
            return str(int(res))
        else:
            return '%.2f' % res

    @classmethod
    def match(cls, target: str) -> bool:
        return True


class AvgFunc(Func):

    def __init__(self, excel_scanner: "ExcelScanner"):
        self.excel_scanner = excel_scanner

    def run(self, target: str) -> str:
        target = target.strip()
        target = target[3:]
        target = target[1:-1]
        row, col = target.split(',')
        total = 0.0
        for sheet in Global.sheets:
            _data = self.excel_scanner.read(self.get_row_num(row.strip()), self.get_col_num(col.strip()), sheet)
            total += float(_data)
        retval = total / len(Global.sheets)
        return '%.2f' % retval

    @classmethod
    def match(cls, target: str) -> bool:
        return re.match(r'AVG\(.*\)', target.strip()) is not None


class MaxFunc(Func):

    def __init__(self, excel_scanner: "ExcelScanner"):
        self.excel_scanner = excel_scanner

    def run(self, target: str) -> str:
        target = target.strip()
        target = target[3:]
        target = target[1:-1]
        row, col = target.split(',')
        _max = 0.0
        for sheet in Global.sheets:
            _data = self.excel_scanner.read(self.get_row_num(row.strip()), self.get_col_num(col.strip()), sheet)
            _max = max(_max, _data)
        return '%.2f' % _max

    @classmethod
    def match(cls, target: str) -> bool:
        return re.match(r'MAX\(.*\)', target.strip()) is not None


class MinFunc(Func):

    def __init__(self, excel_scanner: "ExcelScanner"):
        self.excel_scanner = excel_scanner

    def run(self, target: str) -> str:
        target = target.strip()
        target = target[3:]
        target = target[1:-1]
        row, col = target.split(',')
        _min = 0.0
        for sheet in Global.sheets:
            _data = self.excel_scanner.read(self.get_row_num(row.strip()), self.get_col_num(col.strip()), sheet)
            _min = min(_min, _data)
        return '%.2f' % _min

    @classmethod
    def match(cls, target: str) -> bool:
        return re.match(r'MIN\(.*\)', target.strip()) is not None




