# === IMPORTS ==================================================================

import os
import string
import typing

from enum import Enum
from input import CSV
from xlsxwriter import Workbook

 
# === CLASSES ==================================================================

class XLSX:
    _FILE_NAME : str = 'shotDetect'
    _EXTENSION : str = '.xlsx'
    _SHEET_NAME_RADIUS : str = 'preFloatRadius'
    _SHEET_NAME_LEVEL : str = 'level'
    _HEADER_TIME : str = 'time(s)'
    
    _STATS_LABELS : typing.Tuple[str] = (
        _HEADER_TIME,
        'COUNT',
        'MEAN',
        'STDEV',
        'VARIANCE',
        'MEDIAN',
        'MIN',
        'MAX',
        _HEADER_TIME,
    )
    
    _STATS_FORMULAS : typing.Tuple[str] = (
        'COUNT',
        'AVERAGE',
        'STDEV',
        'VAR',
        'MEDIAN',
        'MIN',
        'MAX',
    )
    
    _ROW_DATA : int = len(_STATS_LABELS)
    
    _ROW_HEADER : typing.Tuple[int] = (
        _STATS_LABELS.index(_HEADER_TIME),
        _STATS_LABELS.index(_HEADER_TIME, 1)
    )
    _ROW_STATS : int = _ROW_HEADER[0] + 1
    
    _COL_TIME : int = 0
    _COL_DATA : int = 1
    
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
    
    def __init__(self):
        self.size : int = 0
        self.wb : Workbook = Workbook(self._FILE_NAME + self._EXTENSION)
        self.radiusSheet : Workbook.worksheet_class = self.wb.add_worksheet(self._SHEET_NAME_RADIUS)
        self.levelSheet : Workbook.worksheet_class = self.wb.add_worksheet(self._SHEET_NAME_LEVEL)
        self._initPreFloatRadiusSheet()
        self._initLevelSheet()
        
    def _initPreFloatRadiusSheet(self):
        for row, label in enumerate(self._STATS_LABELS):
            self.radiusSheet.write(row, 0, label)
    
    def _initLevelSheet(self):
        for row, label in enumerate(self._STATS_LABELS):
            self.levelSheet.write(row, 0, label)
            
    def _addColStats(self, dataSize : int):
        _FORMULAT_FORMAT : str = '={0}({1}{2}:{1}{3})'
        row : int = self._ROW_STATS
        rowStart : int = self._ROW_DATA + 1
        rowEnd : int = rowStart + dataSize
        col : int = self._COL_DATA + self.size
        colStr : str = XLSX._getXLSXColStr(col)
        for formula in self._STATS_FORMULAS:
            self.radiusSheet.write_formula(row, col, _FORMULAT_FORMAT.format(formula, colStr, rowStart, rowEnd))
            self.levelSheet.write_formula(row, col, _FORMULAT_FORMAT.format(formula, colStr, rowStart, rowEnd))
            row += 1
    
    def writeData(self, input : CSV):
        if (len(input.entries) > 0):
            
            # Add the name of the file to the sheets.
            for row in self._ROW_HEADER:
                self.radiusSheet.write(row, self._COL_DATA + self.size, input.name)
                self.levelSheet.write(row, self._COL_DATA + self.size, input.name)
                
            # Add the data to the sheets.
            dataSize = 0
            for i, entry in enumerate(input.entries):
                
                # Only add the data from the shot and the first second before
                # the shot (don't include the 2nd or 3rd second before).
                if entry[2] > 1:
                    break
                
                level = entry[5]
                radius = input.preFloatRadius[i]
                if self.size <= 0:
                    time = entry[1]
                    self.radiusSheet.write_number(self._ROW_DATA + dataSize, self._COL_TIME, time)
                    self.levelSheet.write_number(self._ROW_DATA + dataSize, self._COL_TIME, time)
                self.radiusSheet.write_number(self._ROW_DATA + dataSize, self._COL_DATA + self.size, radius)
                self.levelSheet.write_number(self._ROW_DATA + dataSize, self._COL_DATA + self.size, level)
                dataSize += 1
            self._addColStats(dataSize)
            self.size += 1
    
    def finalize(self):
        self.wb.close()