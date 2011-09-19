# -*- coding: utf-8 -*-
# cython: profile=True

from struct import unpack, pack
import BIFFRecords


cdef class StrCell:

    def __init__(self, rowx, colx, xf_idx, sst_idx):
        self.rowx = rowx
        self.colx = colx
        self.xf_idx = xf_idx
        self.sst_idx = sst_idx

    cpdef get_biff_data(self):
        # return BIFFRecords.LabelSSTRecord(self.rowx, self.colx, self.xf_idx, self.sst_idx).get()
        return pack('<5HL', 0x00FD, 10, self.rowx, self.colx, self.xf_idx, self.sst_idx)

    cpdef get_copy(self):
        cdef StrCell cell = StrCell.__new__(StrCell)
        cell.rowx = self.rowx
        cell.colx = self.colx
        cell.xf_idx = self.xf_idx
        cell.sst_idx = self.sst_idx
        return cell


cdef class BlankCell:

    def __init__(self, rowx, colx, xf_idx):
        self.rowx = rowx
        self.colx = colx
        self.xf_idx = xf_idx

    cpdef get_biff_data(self):
        # return BIFFRecords.BlankRecord(self.rowx, self.colx, self.xf_idx).get()
        return pack('<5H', 0x0201, 6, self.rowx, self.colx, self.xf_idx)

    cpdef get_copy(self):
        #return BlankCell(self.rowx, self.colx, self.xf_idx)
        cdef BlankCell cell = BlankCell.__new__(BlankCell)
        cell.rowx = self.rowx
        cell.colx = self.colx
        cell.xf_idx = self.xf_idx
        return cell


cdef class MulBlankCell:

    def __init__(self, rowx, colx1, colx2, xf_idx):
        self.rowx = rowx
        self.colx1 = colx1
        self.colx2 = colx2
        self.xf_idx = xf_idx

    cpdef get_biff_data(self):
        return BIFFRecords.MulBlankRecord(self.rowx,
            self.colx1, self.colx2, self.xf_idx).get()

    cpdef get_copy(self):
        cdef MulBlankCell cell = MulBlankCell(self.rowx, self.colx1, self.colx2, self.xf_idx)
        return cell


cdef class NumberCell(object):

    def __init__(self, rowx, colx, xf_idx, number):
        self.rowx = rowx
        self.colx = colx
        self.xf_idx = xf_idx
        self.number = float(number)

    cpdef get_encoded_data(self):
        rk_encoded = 0
        num = self.number

        # The four possible kinds of RK encoding are *not* mutually exclusive.
        # The 30-bit integer variety picks up the most.
        # In the code below, the four varieties are checked in descending order
        # of bangs per buck, or not at all.
        # SJM 2007-10-01

        if -0x20000000 <= num < 0x20000000: # fits in 30-bit *signed* int
            inum = int(num)
            if inum == num: # survives round-trip
                # print "30-bit integer RK", inum, hex(inum)
                rk_encoded = 2 | (inum << 2)
                return 1, rk_encoded

        temp = num * 100

        if -0x20000000 <= temp < 0x20000000:
            # That was step 1: the coded value will fit in
            # a 30-bit signed integer.
            itemp = int(round(temp, 0))
            # That was step 2: "itemp" is the best candidate coded value.
            # Now for step 3: simulate the decoding,
            # to check for round-trip correctness.
            if itemp / 100.0 == num:
                # print "30-bit integer RK*100", itemp, hex(itemp)
                rk_encoded = 3 | (itemp << 2)
                return 1, rk_encoded

        if 0: # Cost of extra pack+unpack not justified by tiny yield.
            packed = pack('<d', num)
            w01, w23 = unpack('<2i', packed)
            if not w01 and not(w23 & 3):
                # 34 lsb are 0
                # print "float RK", w23, hex(w23)
                return 1, w23

            packed100 = pack('<d', temp)
            w01, w23 = unpack('<2i', packed100)
            if not w01 and not(w23 & 3):
                # 34 lsb are 0
                # print "float RK*100", w23, hex(w23)
                return 1, w23 | 1

        #print "Number"
        #print
        return 0, pack('<5Hd', 0x0203, 14, self.rowx, self.colx, self.xf_idx, num)

    cpdef get_biff_data(self):
        isRK, value = self.get_encoded_data()
        if isRK:
            return pack('<5Hi', 0x27E, 10, self.rowx, self.colx, self.xf_idx, value)
        return value # NUMBER record already packed

    cpdef get_copy(self):
        cdef NumberCell cell = NumberCell(self.rowx, self.colx, self.xf_idx, self.number)
        return cell

cdef class BooleanCell:

    def __init__(self, rowx, colx, xf_idx, number):
        self.rowx = rowx
        self.colx = colx
        self.xf_idx = xf_idx
        self.number = number

    cpdef get_biff_data(self):
        return BIFFRecords.BoolErrRecord(self.rowx,
            self.colx, self.xf_idx, self.number, 0).get()

    cpdef get_copy(self):
        cdef BooleanCell cell = BooleanCell(self.rowx, self.colx, self.xf_idx, self.number)
        return cell

error_code_map = {
    0x00:  0, # Intersection of two cell ranges is empty
    0x07:  7, # Division by zero
    0x0F: 15, # Wrong type of operand
    0x17: 23, # Illegal or deleted cell reference
    0x1D: 29, # Wrong function or range name
    0x24: 36, # Value range overflow
    0x2A: 42, # Argument or function not available
    '#NULL!' :  0, # Intersection of two cell ranges is empty
    '#DIV/0!':  7, # Division by zero
    '#VALUE!': 36, # Wrong type of operand
    '#REF!'  : 23, # Illegal or deleted cell reference
    '#NAME?' : 29, # Wrong function or range name
    '#NUM!'  : 36, # Value range overflow
    '#N/A!'  : 42, # Argument or function not available
}


cdef class ErrorCell:
    #__slots__ = ["rowx", "colx", "xf_idx", "number", 'error_string_or_code']

    def __init__(self, rowx, colx, xf_idx, error_string_or_code):
        self.rowx = rowx
        self.colx = colx
        self.xf_idx = xf_idx
        self.error_string_or_code = error_string_or_code
        try:
            self.number = error_code_map[error_string_or_code]
        except KeyError:
            raise Exception('Illegal error value (%r)' % error_string_or_code)

    cpdef get_biff_data(self):
        return BIFFRecords.BoolErrRecord(self.rowx,
            self.colx, self.xf_idx, self.number, 1).get()

    cpdef get_copy(self):
        cdef ErrorCell cell = ErrorCell(self.rowx, self.colx, self.xf_idx, self.error_string_or_code)
        return cell


cdef class FormulaCell:
    #__slots__ = ["rowx", "colx", "xf_idx", "frmla", "calc_flags"]

    def __init__(self, rowx, colx, xf_idx, frmla, calc_flags=0):
        self.rowx = rowx
        self.colx = colx
        self.xf_idx = xf_idx
        self.frmla = frmla
        self.calc_flags = calc_flags

    cpdef get_biff_data(self):
        return BIFFRecords.FormulaRecord(self.rowx,
            self.colx, self.xf_idx, self.frmla.rpn(), self.calc_flags).get()

    cpdef get_copy(self):
        cdef FormulaCell cell = FormulaCell(self.rowx, self.colx, self.xf_idx, self.frmla, self.calc_flags)
        return cell


