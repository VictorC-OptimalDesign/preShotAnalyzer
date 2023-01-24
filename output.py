# === IMPORTS ==================================================================

import os
import string
import typing

from enum import Enum
from input import CSV
from xlsxwriter import Workbook

 
# === CLASSES ==================================================================

class XLSX:
    _EXTENSION : str = '.xlsx'
    
    def _getXLSXColStr(col : int) -> str:
        _NUM_LETTERS = len(string.ascii_uppercase)
        pre : int = int(col / _NUM_LETTERS)
        post : int = int(col % _NUM_LETTERS)
        preChar : str = ''
        if (pre > _NUM_LETTERS):
            pre = _NUM_LETTERS
        if (pre > 0):
            preChar = string.ascii_uppercase[pre - 1]
        postChar : str = string.ascii_uppercase[post]
        return preChar + postChar
    
    def __init__(self, fileName : str):
        self.fileName : str = fileName
        self.wb : Workbook = Workbook(self.fileName + self._EXTENSION)
        self.preFloatSheet : Workbook.worksheet_class
        self.levelSheet : Workbook.worksheet_class
        self._initPreFloatSheet()
        self._initLevelSheet()
        
    def _initPreFloatSheet(self):
        pass
    
    def _initLevelSheet(self):
        pass
    
    def writeData(self, input : CSV):
        pass
    
    def finalize(self):
        self.wb.close()