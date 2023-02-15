# === IMPORTS ==================================================================

import math
import os
import typing

from enum import Enum


# === CLASSES ==================================================================

class CSV:
    _EXTENSION : str = '.csv'
    _HEADER_FIELDS : typing.Tuple[str] = (
        'index',
        'time(s)',
        'second',
        'preFloatX',
        'preFloatY',
        'level'
    )
    
    def __init__(self, fileName : str):
        self.fileName : str = fileName
        self.filePath : str = None
        self.name : str = None
        self.header: typing.List[str] = []
        self.entries : typing.List[typing.List[float]] = []
        self.preFloatRadius : typing.List[float] = []
        self.preFloatAngle : typing.List[float] = []
        self._read()
        
    def _read(self):
        if (self.fileName):
            baseName : str = os.path.basename(self.fileName)
            self.name = baseName.replace(self._EXTENSION, '')
            self.filePath = os.path.join(os.getcwd(), self.fileName)
            with open(self.filePath, 'r') as file:
                readLines = file.readlines()
            file.close
            
            numberOfLines : int = len(readLines)
            for i, line in enumerate(readLines):
                fields : typing.List[str] = [x.strip() for x in line.split(',')]
                fields = [i for i in fields if i]
                self._processFields(fields)
                
    def _isHeader(self, fields : typing.List[str]) -> bool:
        result : bool = False
        length : int = len(fields)
        if len(fields) == len(self._HEADER_FIELDS):
            result = True
            for i, field in enumerate(fields):
                if field != self._HEADER_FIELDS[i]:
                    result = False
                    break
        return result
                
    def _processFields(self, fields : typing.List[str]):
        if len(self.header) <= 0 and self._isHeader(fields):
            self.header = fields
        elif len(fields) == len(self._HEADER_FIELDS):
            try:
                values : typing.List[float] = []
                for field in fields:
                    values.append(float(field))
                x : float = values[self._HEADER_FIELDS.index('preFloatX')]
                y : float = values[self._HEADER_FIELDS.index('preFloatY')]
                radius : float = math.sqrt(x**2 + y**2)
                angle : float = math.degrees(math.atan2(y, x))
                self.entries.append(values)
                self.preFloatRadius.append(radius)
                self.preFloatAngle.append(angle)
            except:
                pass
            