# -*- coding: utf-8 -*-
# cython: profile=True


from struct import unpack, pack
import biff_records as BIFFRecords
import Style
cimport cell
import cell
from cell cimport Cell, StrCell, BlankCell, NumberCell, FormulaCell, MulBlankCell, BooleanCell, ErrorCell, \
    _get_cells_biff_data_mul
import ExcelFormula
import datetime as dt

"""
try:
    from decimal import Decimal
except ImportError:
    # Python 2.3: decimal not supported; create dummy Decimal class
    class Decimal(object):
        pass
"""

from decimal import Decimal


cdef class Row:

    cdef public int _idx, _min_col_idx, _max_col_idx, _xf_index, _has_default_xf_index, _height_in_pixels
    cdef public int height, has_default_height, height_mismatch, level, collapse, hidden, space_above, space_below
    cdef public __parent, __parent_wb
    cdef public dict _cells

    cpdef get_book(self)

    cpdef get_parent(self)

    cpdef __adjust_height(self, style)


    cpdef __excel_date_dt(self, date)

    cpdef set_style(self, style)

    cpdef get_cells_count(self)

    cpdef get_row_biff_data(self)

    cpdef insert_cell(self, int col_index, Cell cell_obj)

    cpdef insert_mulcells(self, int colx1, int colx2, Cell cell_obj)

    cpdef get_cells_biff_data(self)

    cpdef get_index(self)

    cpdef set_cell_text(self, int colx, value, style=*)

    cpdef set_cell_blank(self, int colx, style=*)

    cpdef set_cell_mulblanks(self, int first_colx, int last_colx, style=*)

    cpdef write_blanks(self, int first_colx, int last_colx, style=*)

    cpdef set_cell_number(self, int colx, number, style=*)

    cpdef set_cell_date(self, int colx, datetime_obj, style=*)

    cpdef set_cell_formula(self, int colx, formula, style=*, calc_flags=*)

    cpdef set_cell_boolean(self, int colx, value, style=*)

    cpdef set_cell_error(self, int colx, error_string_or_code, style=*)

    cpdef write(self, int col, label, style=*)


    cpdef get_cells(self)

    cpdef _append_cell(self, k, v)

    cpdef move_to(self, int new_idx)

    cpdef get_copy(self, int rowx, parent_sheet)

