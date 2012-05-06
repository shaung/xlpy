# -*- coding: utf-8 -*-
# cython: profile=True


from struct import unpack, pack
import biff_records as BIFFRecords
import style as Style
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

    def __init__(self, rowx, parent_sheet):
        if not (isinstance(rowx, int) and 0 <= rowx <= 65535):
            raise ValueError("row index (%r) not an int in range(65536)" % rowx)
        self._idx = rowx
        self.__parent = parent_sheet
        self.__parent_wb = parent_sheet._parent
        self._cells = {}
        self._min_col_idx = 0
        self._max_col_idx = 0
        self._xf_index = 0x0F
        self._has_default_xf_index = 0
        self._height_in_pixels = 0x11

        self.height = 0x00FF
        self.has_default_height = 0x00
        self.height_mismatch = 0
        self.level = 0
        self.collapse = 0
        self.hidden = 0
        self.space_above = 0
        self.space_below = 0

    cpdef get_book(self):
        return self.__parent_wb

    cpdef get_parent(self):
        return self.__parent

    cpdef __adjust_height(self, style):
        twips = style.font.height
        points = float(twips)/20.0
        # Cell height in pixels can be calcuted by following approx. formula:
        # cell height in pixels = font height in points * 83/50 + 2/5
        # It works when screen resolution is 96 dpi
        pix = int(round(points*83.0/50.0 + 2.0/5.0))
        if pix > self._height_in_pixels:
            self._height_in_pixels = pix


    def __adjust_bound_col_idx(self, *args):
        for arg in args:
            iarg = int(arg)
            if not ((0 <= iarg <= 255) and arg == iarg):
                raise ValueError("column index (%r) not an int in range(256)" % arg)
            sheet = self.__parent
            if iarg < self._min_col_idx:
                self._min_col_idx = iarg
            if iarg > self._max_col_idx:
                self._max_col_idx = iarg
            if iarg < sheet.first_used_col:
                sheet.first_used_col = iarg
            if iarg > sheet.last_used_col:
                sheet.last_used_col = iarg

    cpdef __excel_date_dt(self, date):
        if isinstance(date, dt.date) and (not isinstance(date, dt.datetime)):
            epoch = dt.date(1899, 12, 31)
        elif isinstance(date, dt.time):
            date = dt.datetime.combine(dt.datetime(1900, 1, 1), date)
            epoch = dt.datetime(1900, 1, 1, 0, 0, 0)
        else:
            epoch = dt.datetime(1899, 12, 31, 0, 0, 0)
        delta = date - epoch
        xldate = delta.days + float(delta.seconds) / (24*60*60)
        # Add a day for Excel's missing leap day in 1900
        if xldate > 59:
            xldate += 1
        return xldate

    cpdef set_style(self, style):
        self.__adjust_height(style)
        self._xf_index = self.__parent_wb._styles.add(style)
        self._has_default_xf_index = 1

    cpdef get_cells_count(self):
        return len(self._cells)

    cpdef get_row_biff_data(self):
        height_options = (self.height & 0x07FFF)
        height_options |= (self.has_default_height & 0x01) << 15

        options =  (self.level & 0x07) << 0
        options |= (self.collapse & 0x01) << 4
        options |= (self.hidden & 0x01) << 5
        options |= (self.height_mismatch & 0x01) << 6
        options |= (self._has_default_xf_index & 0x01) << 7
        options |= (0x01 & 0x01) << 8
        options |= (self._xf_index & 0x0FFF) << 16
        options |= (self.space_above & 1) << 28
        options |= (self.space_below & 1) << 29

        return BIFFRecords.RowRecord(self._idx, self._min_col_idx,
            self._max_col_idx, height_options, options).get()

    cpdef insert_cell(self, int col_index, Cell cell_obj):
        cdef Cell prev_cell_obj
        if col_index in self._cells:
            if not self.__parent._cell_overwrite_ok:
                msg = "Attempt to overwrite cell: sheetname=%r rowx=%d colx=%d" \
                    % (self.__parent.name, self._idx, col_index)
                raise Exception(msg)
            prev_cell_obj = self._cells[col_index]
            sst_idx = getattr(prev_cell_obj, 'sst_idx', None)
            if sst_idx is not None:
                self.__parent_wb.del_str(sst_idx)
        self._cells[col_index] = cell_obj

    cpdef insert_mulcells(self, int colx1, int colx2, Cell cell_obj):
        self.insert_cell(colx1, cell_obj)
        for col_index in xrange(colx1+1, colx2+1):
            self.insert_cell(col_index, None)

    cpdef get_cells_biff_data(self):
        cell_items = [item for item in self._cells.iteritems() if item[1] is not None]
        cell_items.sort() # in column order
        return _get_cells_biff_data_mul(self._idx, cell_items)
        # previously:
        # return ''.join([cell.get_biff_data() for colx, cell in cell_items])

    cpdef get_index(self):
        return self._idx

    cpdef set_cell_text(self, int colx, value, style=Style.default_style):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.insert_cell(colx, StrCell(self._idx, colx, xf_index, self.__parent_wb.add_str(value)))

    cpdef set_cell_blank(self, int colx, style=Style.default_style):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.insert_cell(colx, BlankCell(self._idx, colx, xf_index))

    cpdef set_cell_mulblanks(self, int first_colx, int last_colx, style=Style.default_style):
        assert 0 <= first_colx <= last_colx <= 255
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(first_colx, last_colx)
        xf_index = self.__parent_wb._styles.add(style)
        # ncols = last_colx - first_colx + 1
        self.insert_mulcells(first_colx, last_colx, MulBlankCell(self._idx, first_colx, last_colx, xf_index))

    #write_blanks = set_cell_mulblanks
    cpdef write_blanks(self, int first_colx, int last_colx, style=Style.default_style):
        self.set_cell_mulblanks(first_colx, last_colx, style)

    cpdef set_cell_number(self, int colx, number, style=Style.default_style):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.insert_cell(colx, NumberCell(self._idx, colx, xf_index, number))

    cpdef set_cell_date(self, int colx, datetime_obj, style=Style.default_style):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.insert_cell(colx,
            NumberCell(self._idx, colx, xf_index, self.__excel_date_dt(datetime_obj)))

    cpdef set_cell_formula(self, int colx, formula, style=Style.default_style, calc_flags=0):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.__parent_wb.add_sheet_reference(formula)
        self.insert_cell(colx, FormulaCell(self._idx, colx, xf_index, formula, calc_flags=0))

    cpdef set_cell_boolean(self, int colx, value, style=Style.default_style):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.insert_cell(colx, BooleanCell(self._idx, colx, xf_index, bool(value)))

    cpdef set_cell_error(self, int colx, error_string_or_code, style=Style.default_style):
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(colx)
        xf_index = self.__parent_wb._styles.add(style)
        self.insert_cell(colx, ErrorCell(self._idx, colx, xf_index, error_string_or_code))

    cpdef write(self, int col, label, style=None):
        cdef Cell current_cell
        self.__adjust_height(style)
        self.__adjust_bound_col_idx(col)
        if style is None or style == Style.default_style:
            style_index = 0x0F
        else:
            style_index = self.__parent_wb._styles.add(style)

        update_kws = {'xf_idx': style_index}
        if isinstance(label, basestring):
            if len(label) > 0:
                cell_cls = StrCell
                update_kws['sst_idx'] = self.__parent_wb.add_str(label)
            else:
                cell_cls = BlankCell
        elif isinstance(label, bool):
            cell_cls = BooleanCell
            update_kws['number'] = label
        elif isinstance(label, (float, int, long, Decimal)):
            cell_cls = NumberCell
            update_kws['number'] = float(label)
        elif isinstance(label, (dt.datetime, dt.date, dt.time)):
            date_number = self.__excel_date_dt(label)
            cell_cls = NumberCell
            update_kws['number'] = float(date_number)
        elif label is None:
            cell_cls = BlankCell
        elif isinstance(label, ExcelFormula.Formula):
            self.__parent_wb.add_sheet_reference(label)
            cell_cls = FormulaCell
            update_kws['frmla'] = label

        current_cell = self._cells.get(col, None)
        if current_cell:
            if isinstance(current_cell, cell_cls):
                # same type, just update, no new insert
                for k, v in update_kws.iteritems():
                    setattr(current_cell, k, v)
                return

        if isinstance(label, basestring):
            if len(label) > 0:
                self.insert_cell(col,
                    StrCell(self._idx, col, style_index, self.__parent_wb.add_str(label))
                    )
            else:
                self.insert_cell(col, BlankCell(self._idx, col, style_index))
        elif isinstance(label, bool): # bool is subclass of int; test bool first
            self.insert_cell(col, BooleanCell(self._idx, col, style_index, label))
        elif isinstance(label, (float, int, long, Decimal)):
            self.insert_cell(col, NumberCell(self._idx, col, style_index, label))
        elif isinstance(label, (dt.datetime, dt.date, dt.time)):
            date_number = self.__excel_date_dt(label)
            self.insert_cell(col, NumberCell(self._idx, col, style_index, date_number))
        elif label is None:
            self.insert_cell(col, BlankCell(self._idx, col, style_index))
        elif isinstance(label, ExcelFormula.Formula):
            self.__parent_wb.add_sheet_reference(label)
            self.insert_cell(col, FormulaCell(self._idx, col, style_index, label))
        else:
            raise Exception("Unexpected data type %r" % type(label))



    cpdef get_cells(self):
        return self._cells

    cpdef _append_cell(self, k, v):
        self._cells[k] = v

    cpdef move_to(self, int new_idx):
        cdef Cell cell
        self._idx = new_idx
        for indx, cell in self._cells.items():
            if cell is not None:
                try:
                    cell.rowx = new_idx
                except:
                    pass

    cpdef get_copy(self, int rowx, parent_sheet):
        cdef Row row
        cdef Cell cell
        #print 'copy row...%s' % rowx
        other_book = parent_sheet._parent
        same_book = other_book == self.__parent_wb
        row = Row(rowx, parent_sheet)

        for indx in self._cells:
            cell = self._cells[indx]
            if not cell:
                #row._append_cell(indx, cell)
                row._cells[indx] = cell
                continue

            rslt = cell.get_copy()
            if not same_book:
                if cell.xf_idx == 0x0F:
                    rslt.xf_idx = 0x0F
                else:
                    style = self.__parent_wb._styles.get_style(cell.xf_idx)
                    if style:
                        rslt.xf_idx = other_book._styles.add(style)
                    else:
                        rslt.xf_idx = 0x0F

            if isinstance(cell, StrCell):
                sst_idx = cell.sst_idx
                if same_book:
                    self.__parent_wb.add_ref(sst_idx)
                    rslt.sst_idx = sst_idx
                else:
                    s = self.__parent_wb.get_str(sst_idx)
                    rslt.sst_idx = other_book.add_str(s)

            #row._append_cell(indx, rslt)
            row._cells[indx] = rslt

        row._min_col_idx = self._min_col_idx
        row._max_col_idx = self._max_col_idx

        if same_book or self._xf_index == 0x0F:
            row._xf_index = self._xf_index
        else:
            style = self.__parent_wb._styles.get_style(self._xf_index)
            if style:
                xf_idx = other_book._styles.add(style)
            else:
                row._xf_index = 0x0F

        row._has_default_xf_index = self._has_default_xf_index
        row._height_in_pixels = self._height_in_pixels

        row.height = self.height
        row.has_default_height = self.has_default_height
        row.height_mismatch = self.height_mismatch
        row.level = self.level
        row.collapse = self.collapse
        row.hidden = self.hidden
        row.space_above = self.space_above
        row.space_below = self.space_below

        return row
