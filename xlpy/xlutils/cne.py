# -*- coding: utf-8 -*-

from xlpy.xlrd import open_workbook
from xlpy.xlwt import *
from copy import copy as copy_book
from utils import get_xlwt_style_list

def create_copy(fpath):
    wt = open_workbook(fpath, formatting_info=True)
    w = copy_book(wt)
    return CneBook(w, fpath)

"""
    Cne stands for Copy and Edit...
    Not a very good name but the best I can think out for now
"""

class CneSheet:
    def __init__(self, w, sheet, oldsheet):
        self.book = w
        self.worksheet = sheet
        self.oldsheet = oldsheet

    def get_merge_range(self, r, c, ref):
        for mrange in self.oldsheet.merged_cells:
            rlo, rhi, clo, chi = mrange
            if ref == rlo and c == clo:
                #print 'old', mrange, 'new', r, r + rhi - rlo - 1, clo, chi - 1
                return r, r + rhi - rlo - 1, clo, chi - 1
        return None

    def should_ignore(self, r, c):
        for rlo, rhi, clo, chi in self.oldsheet.merged_cells:
            if r > rlo and r <= rhi and c > clo and c <= chi:
                #print 'ignored'
                return True
        return False

    def write_value(self, r, c, value, ref=None):
        style = self.get_style(r, c, ref) or Style.default_style
        mrange = self.get_merge_range(r, c, ref)
        if mrange:
            self.worksheet.write_merge(*mrange, label=value, style=style)
        elif self.should_ignore(ref, c):
            return
        else:
            self.worksheet.write(r, c, value, style)

    def write_row(self, r, ref, *cols):
        #print cols
        cols_dict = dict(cols)
        for c, value in sorted(cols_dict.items()):
            self.write_value(r, c, value, ref)

    def insert_row(self, r, ref, *cols):
        self.worksheet.insert_row_before(r)
        self.write_row(r, ref, *cols)

    def get_style(self, r, c, ref=None):
        style = None
        if ref is not None:
            try:
                cell = self.oldsheet.cell(ref, c) 
                style = self.book.style_list[cell.xf_index]
            except:
                #print 'fallback failed'
                style = None

        if not style:
            try:
                cell = self.oldsheet.cell(r, c) 
                style = self.book.style_list[cell.xf_index]
            except:
                #print r, c ,'style lost'
                style = None

        return style

    def set_value(self, r, c, value, fallback=None):
        #print 'set value'
        style = None
        if fallback is not None:
            try:
                cell = self.oldsheet.cell(fallback, c) 
                style = self.book.style_list[cell.xf_index]
            except:
                #print 'fallback failed'
                style = None

        if not style:
            try:
                cell = self.oldsheet.cell(r, c) 
                style = self.book.style_list[cell.xf_index]
            except:
                #print r, c ,'style lost'
                style = None

        if style:
            self.worksheet.write(r, c, value, style)
        else:
            self.worksheet.write(r, c, value)

    def set_formula(self, r, c, formula):
        #print 'set value'
        cell = self.oldsheet.cell(r, c) 
        self.worksheet.write(r, c, Formula(formula), self.book.style_list[cell.xf_index])

    def __getattr__(self, name):
        if name == 'set_value':
            return self.set_value
        else:
            return getattr(self.worksheet, name)

    def append_horz_page_break(self, value):
        self.worksheet.horz_page_breaks.append(value)

class CneBook:
    def __init__(self, book, old_path):
        self.oldbook = open_workbook(old_path, formatting_info=True)
        self.workbook = book
        self.style_list = get_xlwt_style_list(self.oldbook) 

    def get_sheet_count(self):
        #return self.oldbook.nsheets
        return self.workbook.get_sheet_count()

    def get_sheet(self, n):
        sheet = self.workbook.get_sheet(n)
        oldsheet = self.oldbook.sheet_by_index(n)
        return CneSheet(self, sheet, oldsheet)

    def get_original_sheets(self):
        for i in range(self.get_sheet_count()):
            yield self.workbook.get_sheet(i)

    def copy_sheet(self, n, name):
        sheet = self.workbook.copy_sheet(n, name, cell_overwrite_ok=True)
        oldsheet = self.oldbook.sheet_by_index(n)
        return CneSheet(self, sheet, oldsheet)

    def save(self, fpath):
        self.workbook.save(fpath)

    def __getattr__(self, name):
        if name == 'get_sheet':
            return self.set_value
        else:
            return getattr(self.workbook, name)


