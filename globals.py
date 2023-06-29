import os
from typing import List, ClassVar


class Global:
    excel_scanner: ClassVar = None
    sheets: ClassVar[List[str]] = []
    paragraph_font_family: ClassVar[str] = u"仿宋"
    table_font_family: ClassVar[str] = u"仿宋"
    paragraph_font_size: int = 16
    table_font_size: int = 12
    dst_path: ClassVar = os.getcwd()