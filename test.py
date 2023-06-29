from globals import Global
from scanners import ExcelScanner
from writer import WordWriter

if __name__ == '__main__':

    Global.sheets = ['6.5', '6.6', '6.7', '6.8', '6.9', '6.10', '6.11']
    excel = ExcelScanner(r"test/test.xlsx")
    Global.excel_scanner = excel
    word = WordWriter(r"test/test.docx", r"test/result.docx")
    word.scan()