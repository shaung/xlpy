# -*- coding: windows-1252 -*-

from BIFFRecords import ColInfoRecord

class Column(object):
    def __init__(self, colx, parent_sheet):
        if not(isinstance(colx, int) and 0 <= colx <= 255):
            raise ValueError("column index (%r) not an int in range(256)" % colx)
        self._index = colx
        self._parent = parent_sheet
        self._parent_wb = parent_sheet.get_parent()
        self._xf_index = 0x0F

        self.width = 0x0B92
        self.hidden = 0
        self.level = 0
        self.collapse = 0

    def set_xf_index(self, idx):
        self._xf_index = idx

    def get_book(self):
        return self._parent_wb

    def set_style(self, style):
        self._xf_index = self._parent_wb.add_style(style)

    def width_in_pixels(self):
        # *** Approximation ****
        return int(round(self.width * 0.0272 + 0.446, 0))

    def get_biff_record(self):
        options =  (self.hidden & 0x01) << 0
        options |= (self.level & 0x07) << 8
        options |= (self.collapse & 0x01) << 12

        return ColInfoRecord(self._index, self._index, self.width, self._xf_index, options).get()

    def get_copy(self, indx, parent_sheet):
        col = Column(indx, parent_sheet)
        if parent_sheet.parent == self._parent_wb:
            col._xf_index = self._xf_index
        else:
            style = self._parent_wb.get_style(self._xf_index)
            if style:
                xf_idx = col.get_book().add_style(style)
            else:
                xf_idx = 0x0F
            col._xf_index = xf_idx
        col.width = self.width
        col.hidden = self.hidden
        col.level = self.level
        col.collapse = self.collapse
        return col

