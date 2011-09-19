# -*- coding: utf-8 -*-
# cython: profile=True

from struct import unpack, pack
import BIFFRecords


cdef class StrCell:

    cdef public int rowx, colx, xf_idx, sst_idx

    cpdef get_biff_data(self)

    cpdef get_copy(self)


cdef class BlankCell:
    cdef public int rowx, colx, xf_idx

    cpdef get_biff_data(self)

    cpdef get_copy(self)


cdef class MulBlankCell:
    cdef public int rowx, colx1, colx2, xf_idx

    cpdef get_biff_data(self)

    cpdef get_copy(self)


cdef class NumberCell(object):

    cdef public int rowx, colx, xf_idx
    cdef public float number

    cpdef get_encoded_data(self)

    cpdef get_biff_data(self)

    cpdef get_copy(self)


cdef class BooleanCell:

    cdef public int rowx, colx, xf_idx
    cdef public float number

    cpdef get_biff_data(self)

    cpdef get_copy(self)


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

    cdef public int rowx, colx, xf_idx, number
    cdef public error_string_or_code

    cpdef get_biff_data(self)

    cpdef get_copy(self)


cdef class FormulaCell:
    #__slots__ = ["rowx", "colx", "xf_idx", "frmla", "calc_flags"]
    cdef public int rowx, colx, xf_idx, calc_flags
    cdef public frmla

    cpdef get_biff_data(self)

    cpdef get_copy(self)


# module-level function for *internal* use by the Row module

cdef inline bytes _get_cells_biff_data_mul(int rowx, object cell_items):
    # Return the BIFF data for all cell records in the row.
    # Adjacent BLANK|RK records are combined into MUL(BLANK|RK) records.
    pieces = []
    cdef int nitems = len(cell_items)
    cdef int i = 0
    while i < nitems:
        icolx, icell = cell_items[i]
        if isinstance(icell, NumberCell):
            isRK, value = icell.get_encoded_data()
            if not isRK:
                pieces.append(value) # pre-packed NUMBER record
                i += 1
                continue
            muldata = [(value, icell.xf_idx)]
            target = NumberCell
        elif isinstance(icell, BlankCell):
            muldata = [icell.xf_idx]
            target = BlankCell
        else:
            pieces.append(icell.get_biff_data())
            i += 1
            continue
        lastcolx = icolx
        j = i
        packed_record = ''
        for j in xrange(i+1, nitems):
            jcolx, jcell = cell_items[j]
            if jcolx != lastcolx + 1:
                nexti = j
                break
            if not isinstance(jcell, target):
                nexti = j
                break
            if target == NumberCell:
                isRK, value = jcell.get_encoded_data()
                if not isRK:
                    packed_record = value
                    nexti = j + 1
                    break
                muldata.append((value, jcell.xf_idx))
            else:
                muldata.append(jcell.xf_idx)
            lastcolx = jcolx
        else:
            nexti = j + 1
        if target == NumberCell:
            if lastcolx == icolx:
                # RK record
                value, xf_idx = muldata[0]
                pieces.append(pack('<5Hi', 0x027E, 10, rowx, icolx, xf_idx, value))
            else:
                # MULRK record
                nc = lastcolx - icolx + 1
                pieces.append(pack('<4H', 0x00BD, 6 * nc + 6, rowx, icolx))
                pieces.append(''.join([pack('<Hi', xf_idx, value) for value, xf_idx in muldata]))
                pieces.append(pack('<H', lastcolx))
        else:
            if lastcolx == icolx:
                # BLANK record
                xf_idx = muldata[0]
                pieces.append(pack('<5H', 0x0201, 6, rowx, icolx, xf_idx))
            else:
                # MULBLANK record
                nc = lastcolx - icolx + 1
                pieces.append(pack('<4H', 0x00BE, 2 * nc + 6, rowx, icolx))
                pieces.append(''.join([pack('<H', xf_idx) for xf_idx in muldata]))
                pieces.append(pack('<H', lastcolx))
        if packed_record:
            pieces.append(packed_record)
        i = nexti
    return ''.join(pieces)

