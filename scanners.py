from typing import Optional

import xlrd

from globals import Global


class ExcelScanner:
    xlsx: str = None

    def __init__(self, path: str):
        self.xlsx = xlrd.open_workbook(path)

    def read(self, row: int, col: int, sheet: Optional[str]=None) -> str:
        table = None
        if sheet is None:
            table = self.xlsx.sheet_by_name(Global.sheets[-1])
        else:
            table = self.xlsx.sheet_by_name(sheet)
        return table.cell_value(row, col)

    def get_sheets(self):
        return self.xlsx.sheet_names()



