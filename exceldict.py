#!/usr/bin/env python
import sys

try:
    import openpyxl
except ImportError:
    print 'Error: module openpyxl does not exist.'
    sys.exit(1)


class ExcelDict(object):
    """Convert excel worksheet to python dict"""

    def __init__(self, filename, this_tab):
        self.filename = filename
        self.this_tab = this_tab
        try:
            self.wb = openpyxl.load_workbook(filename)
        except IOError as err:
            print err
            sys.exit(1)

        try:
            self.sheet = self.wb.get_sheet_by_name(this_tab)
        except KeyError as err:
            print err
            sys.exit(1)


    def _getHeader(self):
        header = [ x.value for x in self.sheet[1]]
        return header


    def data(self):
        row = []
        for x in self.sheet[2:self.sheet.max_row]:
            raw_value = [t.value for t in x]
            zipped = dict(zip(self._getHeader(), raw_value))
            row.append(zipped)
        return row

if __name__ == "__main__":
    excel = ExcelDict("file.xlsx", "tab_to_process")
    print excel.data()
