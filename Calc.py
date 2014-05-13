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
    * childType - class of children, if __getitem__ is to be used
    """
    def __init__(self, element):
        assert element.tagName == self.tagName
        self._element = element

    def __getitem__(self, idx):
        children = self._element.getElementsByType(self.childType._odfType)
        if isinstance(idx, tuple):
            if len(idx) > 2:
                return self[idx[0]][idx[1:]]

            return self[idx[0]][idx[1]]

        if isinstance(idx, slice):
            return map(self.childType, children[idx])

        return self.childType(children[idx])


class Cell(_odfObject):
    _odfType = staticmethod(odf.table.TableCell)
    tagName = u'table:table-cell'


class Row(_odfObject):
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


class Sheet(_odfObject):
    tagName = u'table:table'
    childType = Row
    _odfType = staticmethod(odf.table.Table)
    cellPattern = re.compile(r'^([A-Z]+)(\d+)$')

    def __repr__(self):
        return '<Sheet: {}>'.format(self.name)

    @property
    def name(self):
        return self._element.getAttribute("name")

    @name.setter
    def name(self, value):
        if isinstance(value, str):
            value = value.decode('utf-8')
        self._element.setAttribute("name", value)

    @property
    def rows(self):
        return self._element.getElementsByType(odf.table.TableRow)

    @property
    def columns(self):
        return self._element.getElementsByType(odf.table.TableColumn)

    def __getitem__(self, idx):

        # Allow indexing by cell name (e.g. E10)
        # Slicing by cell names not implemented
        if isinstance(idx, str):
            test = self.cellPattern.match(idx)
            if test:
                col, row = test.groups()
                return self[int(row) - 1, col]
        
        return super(Sheet, self).__getitem__(idx)


class Spreadsheet(object):
    def __init__(self, path):
        self.path = path

        if not path.endswith('.ods'):
            self.path = self.path + '.ods'

        if os.path.isfile(self.path):
            self._document = odf.opendocument.load(self.path)
        else:
            self._document = odf.opendocument.OpenDocumentSpreadsheet()

    def save(self):
        self._document.save(self.path)

    @property
    def sheets(self):
        return map(Sheet, self._document.spreadsheet.getElementsByType(
            odf.table.Table))
