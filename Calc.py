#!/usr/bin/env python
#
# 2014 Christopher J Johnson
import os
import re
import odf.opendocument
import odf.table
import odf.text


class _odfObject(object):
    """Abstract class for odf.element.Element wrappers
    
    Subclasses should define:
    * _odfType  - a static method with the element type they wrap
    * tagName   - unicode string containing element tagName to compare against
    """
    def __init__(self, element):
        assert element.tagName == self.tagName
        self._element = element


class _odfIndexable(_odfObject):
    """Abstract class for wrappers with known child types

    Subclasses should define:
    * childType - class of children, if __getitem__ is to be used
    """
    def __getitem__(self, idx):
        children = self._element.getElementsByType(self.childType._odfType)
        if isinstance(idx, tuple):
            sub = idx[1]
            if len(idx) > 2:
                sub = idx[1:]

            if isinstance(idx[0], slice):
                return [self.childType(child)[sub]
                        for child in children[idx[0]]]

            return self[idx[0]][sub]

        if isinstance(idx, slice):
            return map(self.childType, children[idx])

        return self.childType(children[idx])


class Cell(_odfObject):
    _odfType = staticmethod(odf.table.TableCell)
    tagName = u'table:table-cell'


class Row(_odfIndexable):
    childType = Cell
    _odfType = staticmethod(odf.table.TableRow)
    tagName = u'table:table-row'

    def __getitem__(self, idx):
        # Allow indexing by letter
        # Slicing by letter not implemented
        if isinstance(idx, str):
            idx = reduce(lambda x, y: 26 * (x + 1) + y,
                         (ord(x) - ord('A') for x in idx))
        return super(Row, self).__getitem__(idx)


class Table(_odfIndexable):
    tagName = u'table:table'
    childType = Row
    _odfType = staticmethod(odf.table.Table)
    cellPattern = re.compile(r'^([A-Z]+)(\d+)$')

    def __repr__(self):
        return '<Table: {}>'.format(self.name)

    @property
    def name(self):
        return self._element.getAttribute("name")

    @name.setter
    def name(self, value):
        if isinstance(value, str):
            value = value.decode('utf-8')
        self._element.setAttribute("name", value)

    def __getitem__(self, idx):

        # Allow indexing by cell name (e.g. E10)
        # Slicing by cell names not implemented
        if isinstance(idx, str):
            test = self.cellPattern.match(idx)
            if test:
                col, row = test.groups()
                return self[int(row) - 1, col]
        
        return super(Table, self).__getitem__(idx)


class Spreadsheet(_odfIndexable):
    childType = Table

    def __init__(self, path):
        self.path = path

        if not path.endswith('.ods'):
            self.path = self.path + '.ods'

        if os.path.isfile(self.path):
            self._document = odf.opendocument.load(self.path)
        else:
            self._document = odf.opendocument.OpenDocumentSpreadsheet()

        self._element = self._document.spreadsheet

    def __len__(self):
        return len(self.tables)

    def save(self):
        self._document.save(self.path)

    @property
    def tables(self):
        return map(Table, self._document.spreadsheet.getElementsByType(
            odf.table.Table))
