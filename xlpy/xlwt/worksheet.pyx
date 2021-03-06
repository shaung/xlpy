# -*- coding: utf-8 -*-
# cython: profile=True

'''
            BOF
            UNCALCED
            INDEX
            Calculation Settings Block
            PRINTHEADERS
            PRINTGRIDLINES
            GRIDSET
            GUTS
            DEFAULTROWHEIGHT
            WSBOOL
            Page Settings Block
            Worksheet Protection Block
            DEFCOLWIDTH
            COLINFO
            SORT
            DIMENSIONS
            Row Blocks
            WINDOW2
            SCL
            PANE
            SELECTION
            STANDARDWIDTH
            MERGEDCELLS
            LABELRANGES
            PHONETIC
            Conditional Formatting Table
            Hyperlink Table
            Data Validity Table
            SHEETLAYOUT (BIFF8X only)
            SHEETPROTECTION (BIFF8X only)
            RANGEPROTECTION (BIFF8X only)
            EOF
'''

import biff_records as BIFFRecords
import Bitmap
import formatting as Formatting
import style as Style
import os, tempfile
import odraw

from column import Column
import row
from row cimport Row
import cell
from cell cimport Cell


cdef class Worksheet:

    cdef object __weakref__

    cdef Row, Column

    cdef public __name, _parent, _cell_overwrite_ok
    cdef public dict  __rows, __cols
    cdef readonly list __merged_ranges
    cdef public __bmp_rec, __idx, __pic_rec
    cdef public int __show_formulas, __show_grid, __show_headers, __panes_frozen, show_zero_values
    cdef public int __auto_colour_grid, __cols_right_to_left, __show_outline, __remove_splits, __selected
    cdef public int __sheet_visible, __page_preview, __first_visible_row, __first_visible_col, __grid_colour
    cdef public __preview_magn, __normal_magn, __scl_magn, explicit_magn_setting, visibility
    cdef public __vert_split_pos, __horz_split_pos, __vert_split_first_visible, __horz_split_first_visible
    cdef public split_position_units_are_twips
    cdef public __row_gut_width, __col_gut_height
    cdef public __show_auto_page_breaks, __dialogue_sheet, __auto_style_outline, __outline_below, __outline_right
    cdef public __fit_num_pages, __show_row_outline, __show_col_outline, __alt_expr_eval, __alt_formula_entries
    cdef public __row_default_height, row_default_height_mismatch, row_default_hidden, row_default_space_above, row_default_space_below
    cdef public __col_default_width
    cdef public int __calc_mode, __calc_count, __RC_ref_mode, __iterations_on, __save_recalc
    cdef public float __delta
    cdef public int __print_headers, __print_grid, __grid_set
    cdef public __vert_page_breaks, __horz_page_breaks, __header_str, __footer_str, __print_centered_vert, __print_centered_horz
    cdef public float __left_margin, __right_margin, __top_margin, __bottom_margin
    cdef public int __paper_size_code, __print_scaling, __start_page_number, __fit_width_to_pages, __fit_height_to_pages
    cdef public int __print_in_rows, __portrait, __print_not_colour, __print_draft, __print_notes, __print_notes_at_end
    cdef public __print_omit_errors, __print_hres, __print_vres, __header_margin, __footer_margin, __copies_num
    cdef public int __wnd_protect, __obj_protect, __protect, __scen_protect
    cdef public __password
    cdef public int last_used_row, first_used_row, last_used_col, first_used_col
    cdef public row_tempfile, __flushed_rows, __row_visible_levels

    # a safe default value, 3 is always valid!
    active_pane = 3
    
    #################################################################
    ## Constructor
    #################################################################
    def __init__(self, sheetname, parent_book, cell_overwrite_ok=False):

        self.Row = Row
        self.Column = Column

        self.__name = sheetname
        self._parent = parent_book
        self._cell_overwrite_ok = cell_overwrite_ok

        self.__rows = {}
        self.__cols = {}
        self.__merged_ranges = []
        self.__bmp_rec = ''
        self.__idx = parent_book.get_sheet_count()
        self.__pic_rec = odraw.PictureSection(parent_book.get_sheet_count(), self)
        self._parent.drawing_group.add_sheet(self.__idx, len(self.__pic_rec.pics))

        self.__show_formulas = 0
        self.__show_grid = 1
        self.__show_headers = 1
        self.__panes_frozen = 0
        self.show_zero_values = 1
        self.__auto_colour_grid = 1
        self.__cols_right_to_left = 0
        self.__show_outline = 1
        self.__remove_splits = 0
        # Multiple sheets can be selected, but only one can be active
        # (hold down Ctrl and click multiple tabs in the file in OOo)
        self.__selected = 0
        # "sheet_visible" should really be called "sheet_active"
        # and is 1 when this sheet is the sheet displayed when the file
        # is open. More than likely only one sheet should ever be set as
        # visible.
        # The same sheet should be specified in Workbook.active_sheet
        # (that way, both the WINDOW1 record in the book and the WINDOW2
        # records in each sheet will be in agreement)
        # The visibility of the sheet is found in the "visibility"
        # attribute obtained from the BOUNDSHEET record.
        self.__sheet_visible = 0
        self.__page_preview = 0

        self.__first_visible_row = 0
        self.__first_visible_col = 0
        self.__grid_colour = 0x40
        self.__preview_magn = 0 # use default (60%)
        self.__normal_magn = 0 # use default (100%)
        self.__scl_magn = None
        self.explicit_magn_setting = False

        self.visibility = 0 # from/to BOUNDSHEET record.

        self.__vert_split_pos = None
        self.__horz_split_pos = None
        self.__vert_split_first_visible = None
        self.__horz_split_first_visible = None

        # This is a caller-settable flag:

        self.split_position_units_are_twips = False

        # Default is False for backward compatibility with pyExcelerator
        # and previous versions of xlwt.
        #   if panes_frozen:
        #       vert/horz_split_pos are taken as number of rows/cols
        #   else: # split
        #       if split_position_units_are_twips:
        #           vert/horz_split_pos are taken as number of twips
        #       else:
        #           vert/horz_split_pos are taken as
        #           number of rows(cols) * default row(col) height (width) (i.e. 12.75 (8.43) somethings)
        #           and converted to twips by approximate formulas
        # Callers who are copying an existing file should use
        #     xlwt_worksheet.split_position_units_are_twips = True
        # because that's what's actually in the file.

		# There are 20 twips to a point. There are 72 points to an inch.

        self.__row_gut_width = 0
        self.__col_gut_height = 0

        self.__show_auto_page_breaks = 1
        self.__dialogue_sheet = 0
        self.__auto_style_outline = 0
        self.__outline_below = 0
        self.__outline_right = 0
        self.__fit_num_pages = 0
        self.__show_row_outline = 1
        self.__show_col_outline = 1
        self.__alt_expr_eval = 0
        self.__alt_formula_entries = 0

        self.__row_default_height = 0x00FF
        self.row_default_height_mismatch = 0
        self.row_default_hidden = 0
        self.row_default_space_above = 0
        self.row_default_space_below = 0

        self.__col_default_width = 0x0008

        self.__calc_mode = 1
        self.__calc_count = 0x0064
        self.__RC_ref_mode = 1
        self.__iterations_on = 0
        self.__delta = 0.001
        self.__save_recalc = 0

        self.__print_headers = 0
        self.__print_grid = 0
        self.__grid_set = 1
        self.__vert_page_breaks = []
        self.__horz_page_breaks = []
        self.__header_str = '&P'
        self.__footer_str = '&F'
        self.__print_centered_vert = 0
        self.__print_centered_horz = 1
        self.__left_margin = 0.3 #0.5
        self.__right_margin = 0.3 #0.5
        self.__top_margin = 0.61 #1.0
        self.__bottom_margin = 0.37 #1.0
        self.__paper_size_code = 9 # A4
        self.__print_scaling = 100
        self.__start_page_number = 1
        self.__fit_width_to_pages = 1
        self.__fit_height_to_pages = 1
        self.__print_in_rows = 1
        self.__portrait = 1
        self.__print_not_colour = 0
        self.__print_draft = 0
        self.__print_notes = 0
        self.__print_notes_at_end = 0
        self.__print_omit_errors = 0
        self.__print_hres = 0x0258 # 600 dpi
        self.__print_vres = 0x04B0 # 1200 dpi
        self.__header_margin = 0.1
        self.__footer_margin = 0.1
        self.__copies_num = 1

        self.__wnd_protect = 0
        self.__obj_protect = 0
        self.__protect = 0
        self.__scen_protect = 0
        self.__password = ''

        self.last_used_row = 0
        self.first_used_row = 65535
        self.last_used_col = 0
        self.first_used_col = 255
        self.row_tempfile = None
        self.__flushed_rows = {}
        self.__row_visible_levels = 0

    #################################################################
    ## Properties, "getters", "setters"
    #################################################################

    def set_name(self, value):
        self.__name = value

    def get_name(self):
        return self.__name

    name = property(get_name, set_name)

    #################################################################

    def get_parent(self):
        return self._parent

    parent = property(get_parent)

    #################################################################

    def get_rows(self):
        return self.__rows

    rows = property(get_rows)

    #################################################################

    def get_cols(self):
        return self.__cols

    cols = property(get_cols)

    #################################################################

    def get_merged_ranges(self):
        return self.__merged_ranges

    merged_ranges = property(get_merged_ranges)

    #################################################################

    def get_bmp_rec(self):
        return self.__bmp_rec

    bmp_rec = property(get_bmp_rec)

    #################################################################

    def get_pic_rec(self):
        return self.__pic_rec

    pic_rec = property(get_pic_rec)

    #################################################################

    def set_show_formulas(self, value):
        self.__show_formulas = int(value)

    def get_show_formulas(self):
        return bool(self.__show_formulas)

    show_formulas = property(get_show_formulas, set_show_formulas)

    #################################################################

    def set_show_grid(self, value):
        self.__show_grid = int(value)

    def get_show_grid(self):
        return bool(self.__show_grid)

    show_grid = property(get_show_grid, set_show_grid)

    #################################################################

    def set_show_headers(self, value):
        self.__show_headers = int(value)

    def get_show_headers(self):
        return bool(self.__show_headers)

    show_headers = property(get_show_headers, set_show_headers)

    #################################################################

    def set_panes_frozen(self, value):
        self.__panes_frozen = int(value)

    def get_panes_frozen(self):
        return bool(self.__panes_frozen)

    panes_frozen = property(get_panes_frozen, set_panes_frozen)

    #################################################################

    ### def set_show_empty_as_zero(self, value):
    ###     self.__show_empty_as_zero = int(value)

    ### def get_show_empty_as_zero(self):
    ###     return bool(self.__show_empty_as_zero)

    ### show_empty_as_zero = property(get_show_empty_as_zero, set_show_empty_as_zero)

    #################################################################

    def set_auto_colour_grid(self, value):
        self.__auto_colour_grid = int(value)

    def get_auto_colour_grid(self):
        return bool(self.__auto_colour_grid)

    auto_colour_grid = property(get_auto_colour_grid, set_auto_colour_grid)

    #################################################################

    def set_cols_right_to_left(self, value):
        self.__cols_right_to_left = int(value)

    def get_cols_right_to_left(self):
        return bool(self.__cols_right_to_left)

    cols_right_to_left = property(get_cols_right_to_left, set_cols_right_to_left)

    #################################################################

    def set_show_outline(self, value):
        self.__show_outline = int(value)

    def get_show_outline(self):
        return bool(self.__show_outline)

    show_outline = property(get_show_outline, set_show_outline)

    #################################################################

    def set_remove_splits(self, value):
        self.__remove_splits = int(value)

    def get_remove_splits(self):
        return bool(self.__remove_splits)

    remove_splits = property(get_remove_splits, set_remove_splits)

    #################################################################

    def set_selected(self, value):
        self.__selected = int(value)

    def get_selected(self):
        return bool(self.__selected)

    selected = property(get_selected, set_selected)

    #################################################################

    def set_sheet_visible(self, value):
        self.__sheet_visible = int(value)

    def get_sheet_visible(self):
        return bool(self.__sheet_visible)

    sheet_visible = property(get_sheet_visible, set_sheet_visible)

    #################################################################

    def set_page_preview(self, value):
        self.__page_preview = int(value)

    def get_page_preview(self):
        return bool(self.__page_preview)

    page_preview = property(get_page_preview, set_page_preview)

    #################################################################

    def set_first_visible_row(self, value):
        self.__first_visible_row = value

    def get_first_visible_row(self):
        return self.__first_visible_row

    first_visible_row = property(get_first_visible_row, set_first_visible_row)

    #################################################################

    def set_first_visible_col(self, value):
        self.__first_visible_col = value

    def get_first_visible_col(self):
        return self.__first_visible_col

    first_visible_col = property(get_first_visible_col, set_first_visible_col)

    #################################################################

    def set_grid_colour(self, value):
        self.__grid_colour = value

    def get_grid_colour(self):
        return self.__grid_colour

    grid_colour = property(get_grid_colour, set_grid_colour)

    #################################################################

    def set_preview_magn(self, value):
        self.__preview_magn = value

    def get_preview_magn(self):
        return self.__preview_magn

    preview_magn = property(get_preview_magn, set_preview_magn)

    #################################################################

    def set_normal_magn(self, value):
        self.__normal_magn = value

    def get_normal_magn(self):
        return self.__normal_magn

    normal_magn = property(get_normal_magn, set_normal_magn)

    #################################################################

    def set_scl_magn(self, value):
        self.__scl_magn = value

    def get_scl_magn(self):
        return self.__scl_magn

    scl_magn = property(get_scl_magn, set_scl_magn)


    #################################################################

    def set_vert_split_pos(self, value):
        self.__vert_split_pos = abs(value)

    def get_vert_split_pos(self):
        return self.__vert_split_pos

    vert_split_pos = property(get_vert_split_pos, set_vert_split_pos)

    #################################################################

    def set_horz_split_pos(self, value):
        self.__horz_split_pos = abs(value)

    def get_horz_split_pos(self):
        return self.__horz_split_pos

    horz_split_pos = property(get_horz_split_pos, set_horz_split_pos)

    #################################################################

    def set_vert_split_first_visible(self, value):
        self.__vert_split_first_visible = abs(value)

    def get_vert_split_first_visible(self):
        return self.__vert_split_first_visible

    vert_split_first_visible = property(get_vert_split_first_visible, set_vert_split_first_visible)

    #################################################################

    def set_horz_split_first_visible(self, value):
        self.__horz_split_first_visible = abs(value)

    def get_horz_split_first_visible(self):
        return self.__horz_split_first_visible

    horz_split_first_visible = property(get_horz_split_first_visible, set_horz_split_first_visible)

    #################################################################

    #def set_row_gut_width(self, value):
    #    self.__row_gut_width = value
    #
    #def get_row_gut_width(self):
    #    return self.__row_gut_width
    #
    #row_gut_width = property(get_row_gut_width, set_row_gut_width)
    #
    #################################################################
    #
    #def set_col_gut_height(self, value):
    #    self.__col_gut_height = value
    #
    #def get_col_gut_height(self):
    #    return self.__col_gut_height
    #
    #col_gut_height = property(get_col_gut_height, set_col_gut_height)
    #
    #################################################################

    def set_show_auto_page_breaks(self, value):
        self.__show_auto_page_breaks = int(value)

    def get_show_auto_page_breaks(self):
        return bool(self.__show_auto_page_breaks)

    show_auto_page_breaks = property(get_show_auto_page_breaks, set_show_auto_page_breaks)

    #################################################################

    def set_dialogue_sheet(self, value):
        self.__dialogue_sheet = int(value)

    def get_dialogue_sheet(self):
        return bool(self.__dialogue_sheet)

    dialogue_sheet = property(get_dialogue_sheet, set_dialogue_sheet)

    #################################################################

    def set_auto_style_outline(self, value):
        self.__auto_style_outline = int(value)

    def get_auto_style_outline(self):
        return bool(self.__auto_style_outline)

    auto_style_outline = property(get_auto_style_outline, set_auto_style_outline)

    #################################################################

    def set_outline_below(self, value):
        self.__outline_below = int(value)

    def get_outline_below(self):
        return bool(self.__outline_below)

    outline_below = property(get_outline_below, set_outline_below)

    #################################################################

    def set_outline_right(self, value):
        self.__outline_right = int(value)

    def get_outline_right(self):
        return bool(self.__outline_right)

    outline_right = property(get_outline_right, set_outline_right)

    #################################################################

    def set_fit_num_pages(self, value):
        self.__fit_num_pages = value

    def get_fit_num_pages(self):
        return self.__fit_num_pages

    fit_num_pages = property(get_fit_num_pages, set_fit_num_pages)

    #################################################################

    def set_show_row_outline(self, value):
        self.__show_row_outline = int(value)

    def get_show_row_outline(self):
        return bool(self.__show_row_outline)

    show_row_outline = property(get_show_row_outline, set_show_row_outline)

    #################################################################

    def set_show_col_outline(self, value):
        self.__show_col_outline = int(value)

    def get_show_col_outline(self):
        return bool(self.__show_col_outline)

    show_col_outline = property(get_show_col_outline, set_show_col_outline)

    #################################################################

    def set_alt_expr_eval(self, value):
        self.__alt_expr_eval = int(value)

    def get_alt_expr_eval(self):
        return bool(self.__alt_expr_eval)

    alt_expr_eval = property(get_alt_expr_eval, set_alt_expr_eval)

    #################################################################

    def set_alt_formula_entries(self, value):
        self.__alt_formula_entries = int(value)

    def get_alt_formula_entries(self):
        return bool(self.__alt_formula_entries)

    alt_formula_entries = property(get_alt_formula_entries, set_alt_formula_entries)

    #################################################################

    def set_row_default_height(self, value):
        self.__row_default_height = value

    def get_row_default_height(self):
        return self.__row_default_height

    row_default_height = property(get_row_default_height, set_row_default_height)

    #################################################################

    def set_col_default_width(self, value):
        self.__col_default_width = value

    def get_col_default_width(self):
        return self.__col_default_width

    col_default_width = property(get_col_default_width, set_col_default_width)

    #################################################################

    def set_calc_mode(self, value):
        self.__calc_mode = value & 0x03

    def get_calc_mode(self):
        return self.__calc_mode

    calc_mode = property(get_calc_mode, set_calc_mode)

    #################################################################

    def set_calc_count(self, value):
        self.__calc_count = value

    def get_calc_count(self):
        return self.__calc_count

    calc_count = property(get_calc_count, set_calc_count)

    #################################################################

    def set_RC_ref_mode(self, value):
        self.__RC_ref_mode = int(value)

    def get_RC_ref_mode(self):
        return bool(self.__RC_ref_mode)

    RC_ref_mode = property(get_RC_ref_mode, set_RC_ref_mode)

    #################################################################

    def set_iterations_on(self, value):
        self.__iterations_on = int(value)

    def get_iterations_on(self):
        return bool(self.__iterations_on)

    iterations_on = property(get_iterations_on, set_iterations_on)

    #################################################################

    def set_delta(self, value):
        self.__delta = value

    def get_delta(self):
        return self.__delta

    delta = property(get_delta, set_delta)

    #################################################################

    def set_save_recalc(self, value):
        self.__save_recalc = int(value)

    def get_save_recalc(self):
        return bool(self.__save_recalc)

    save_recalc = property(get_save_recalc, set_save_recalc)

    #################################################################

    def set_print_headers(self, value):
        self.__print_headers = int(value)

    def get_print_headers(self):
        return bool(self.__print_headers)

    print_headers = property(get_print_headers, set_print_headers)

    #################################################################

    def set_print_grid(self, value):
        self.__print_grid = int(value)

    def get_print_grid(self):
        return bool(self.__print_grid)

    print_grid = property(get_print_grid, set_print_grid)

    #################################################################
    #
    #def set_grid_set(self, value):
    #    self.__grid_set = int(value)
    #
    #def get_grid_set(self):
    #    return bool(self.__grid_set)
    #
    #grid_set = property(get_grid_set, set_grid_set)
    #
    #################################################################

    def set_vert_page_breaks(self, value):
        self.__vert_page_breaks = value

    def get_vert_page_breaks(self):
        return self.__vert_page_breaks

    vert_page_breaks = property(get_vert_page_breaks, set_vert_page_breaks)

    #################################################################

    def set_horz_page_breaks(self, value):
        self.__horz_page_breaks = value

    def get_horz_page_breaks(self):
        return self.__horz_page_breaks

    horz_page_breaks = property(get_horz_page_breaks, set_horz_page_breaks)

    #################################################################

    def set_header_str(self, value):
        if isinstance(value, str):
            value = unicode(value, self._parent.encoding)
        self.__header_str = value

    def get_header_str(self):
        return self.__header_str

    header_str = property(get_header_str, set_header_str)

    #################################################################

    def set_footer_str(self, value):
        if isinstance(value, str):
            value = unicode(value, self._parent.encoding)
        self.__footer_str = value

    def get_footer_str(self):
        return self.__footer_str

    footer_str = property(get_footer_str, set_footer_str)

    #################################################################

    def set_print_centered_vert(self, value):
        self.__print_centered_vert = int(value)

    def get_print_centered_vert(self):
        return bool(self.__print_centered_vert)

    print_centered_vert = property(get_print_centered_vert, set_print_centered_vert)

    #################################################################

    def set_print_centered_horz(self, value):
        self.__print_centered_horz = int(value)

    def get_print_centered_horz(self):
        return bool(self.__print_centered_horz)

    print_centered_horz = property(get_print_centered_horz, set_print_centered_horz)

    #################################################################

    def set_left_margin(self, value):
        self.__left_margin = value

    def get_left_margin(self):
        return self.__left_margin

    left_margin = property(get_left_margin, set_left_margin)

    #################################################################

    def set_right_margin(self, value):
        self.__right_margin = value

    def get_right_margin(self):
        return self.__right_margin

    right_margin = property(get_right_margin, set_right_margin)

    #################################################################

    def set_top_margin(self, value):
        self.__top_margin = value

    def get_top_margin(self):
        return self.__top_margin

    top_margin = property(get_top_margin, set_top_margin)

    #################################################################

    def set_bottom_margin(self, value):
        self.__bottom_margin = value

    def get_bottom_margin(self):
        return self.__bottom_margin

    bottom_margin = property(get_bottom_margin, set_bottom_margin)

    #################################################################

    def set_paper_size_code(self, value):
        self.__paper_size_code = value

    def get_paper_size_code(self):
        return self.__paper_size_code

    paper_size_code = property(get_paper_size_code, set_paper_size_code)

    #################################################################

    def set_print_scaling(self, value):
        self.__print_scaling = value

    def get_print_scaling(self):
        return self.__print_scaling

    print_scaling = property(get_print_scaling, set_print_scaling)

    #################################################################

    def set_start_page_number(self, value):
        self.__start_page_number = value

    def get_start_page_number(self):
        return self.__start_page_number

    start_page_number = property(get_start_page_number, set_start_page_number)

    #################################################################

    def set_fit_width_to_pages(self, value):
        self.__fit_width_to_pages = value

    def get_fit_width_to_pages(self):
        return self.__fit_width_to_pages

    fit_width_to_pages = property(get_fit_width_to_pages, set_fit_width_to_pages)

    #################################################################

    def set_fit_height_to_pages(self, value):
        self.__fit_height_to_pages = value

    def get_fit_height_to_pages(self):
        return self.__fit_height_to_pages

    fit_height_to_pages = property(get_fit_height_to_pages, set_fit_height_to_pages)

    #################################################################

    def set_print_in_rows(self, value):
        self.__print_in_rows = int(value)

    def get_print_in_rows(self):
        return bool(self.__print_in_rows)

    print_in_rows = property(get_print_in_rows, set_print_in_rows)

    #################################################################

    def set_portrait(self, value):
        self.__portrait = int(value)

    def get_portrait(self):
        return bool(self.__portrait)

    portrait = property(get_portrait, set_portrait)

    #################################################################

    def set_print_colour(self, value):
        self.__print_not_colour = int(not value)

    def get_print_colour(self):
        return not bool(self.__print_not_colour)

    print_colour = property(get_print_colour, set_print_colour)

    #################################################################

    def set_print_draft(self, value):
        self.__print_draft = int(value)

    def get_print_draft(self):
        return bool(self.__print_draft)

    print_draft = property(get_print_draft, set_print_draft)

    #################################################################

    def set_print_notes(self, value):
        self.__print_notes = int(value)

    def get_print_notes(self):
        return bool(self.__print_notes)

    print_notes = property(get_print_notes, set_print_notes)

    #################################################################

    def set_print_notes_at_end(self, value):
        self.__print_notes_at_end = int(value)

    def get_print_notes_at_end(self):
        return bool(self.__print_notes_at_end)

    print_notes_at_end = property(get_print_notes_at_end, set_print_notes_at_end)

    #################################################################

    def set_print_omit_errors(self, value):
        self.__print_omit_errors = int(value)

    def get_print_omit_errors(self):
        return bool(self.__print_omit_errors)

    print_omit_errors = property(get_print_omit_errors, set_print_omit_errors)

    #################################################################

    def set_print_hres(self, value):
        self.__print_hres = value

    def get_print_hres(self):
        return self.__print_hres

    print_hres = property(get_print_hres, set_print_hres)

    #################################################################

    def set_print_vres(self, value):
        self.__print_vres = value

    def get_print_vres(self):
        return self.__print_vres

    print_vres = property(get_print_vres, set_print_vres)

    #################################################################

    def set_header_margin(self, value):
        self.__header_margin = value

    def get_header_margin(self):
        return self.__header_margin

    header_margin = property(get_header_margin, set_header_margin)

    #################################################################

    def set_footer_margin(self, value):
        self.__footer_margin = value

    def get_footer_margin(self):
        return self.__footer_margin

    footer_margin = property(get_footer_margin, set_footer_margin)

    #################################################################

    def set_copies_num(self, value):
        self.__copies_num = value

    def get_copies_num(self):
        return self.__copies_num

    copies_num = property(get_copies_num, set_copies_num)

    ##################################################################

    def set_wnd_protect(self, value):
        self.__wnd_protect = int(value)

    def get_wnd_protect(self):
        return bool(self.__wnd_protect)

    wnd_protect = property(get_wnd_protect, set_wnd_protect)

    #################################################################

    def set_obj_protect(self, value):
        self.__obj_protect = int(value)

    def get_obj_protect(self):
        return bool(self.__obj_protect)

    obj_protect = property(get_obj_protect, set_obj_protect)

    #################################################################

    def set_protect(self, value):
        self.__protect = int(value)

    def get_protect(self):
        return bool(self.__protect)

    protect = property(get_protect, set_protect)

    #################################################################

    def set_scen_protect(self, value):
        self.__scen_protect = int(value)

    def get_scen_protect(self):
        return bool(self.__scen_protect)

    scen_protect = property(get_scen_protect, set_scen_protect)

    #################################################################

    def set_password(self, value):
        self.__password = value

    def get_password(self):
        return self.__password

    password = property(get_password, set_password)

    def get_index(self):
        return self.__idx

    index = property(get_index)

    def is_visible(self):
        return self.visibility == 0

    def is_hidden(self):
        return self.visibility == -1
    
    def is_very_hidden(self):
        return self.visibility == -2

    def set_visible(self):
        self.visibility = 0
 
    def set_hidden(self):
        self.visibility = -1
 
    def set_very_hidden(self):
        self.visibility = -2
 
    ##################################################################
    ## Methods
    ##################################################################

    def write(self, r, c, label="", style=Style.default_style):
        self.row(r).write(c, label, style)

    def merge(self, r1, r2, c1, c2, style=Style.default_style):
        # Stand-alone merge of previously written cells.
        # Problems: (1) style to be used should be existing style of
        # the top-left cell, not an arg.
        # (2) should ensure that any previous data value in
        # non-top-left cells is nobbled.
        # Note: if a cell is set by a data record then later
        # is referenced by a [MUL]BLANK record, Excel will blank
        # out the cell on the screen, but OOo & Gnu will not
        # blank it out. Need to do something better than writing
        # multiple records. In the meantime, avoid this method and use
        # write_merge() instead.
        if c2 > c1:
            self.row(r1).write_blanks(c1 + 1, c2,  style)
        for r in range(r1+1, r2+1):
            self.row(r).write_blanks(c1, c2,  style)
        self.__merged_ranges.append((r1, r2, c1, c2))

    def write_merge(self, r1, r2, c1, c2, label="", style=Style.default_style):
        assert 0 <= c1 <= c2 <= 255
        assert 0 <= r1 <= r2 <= 65535
        self.write(r1, c1, label, style)
        if c2 > c1:
            self.row(r1).write_blanks(c1 + 1, c2,  style) # skip (r1, c1)
        for r in range(r1+1, r2+1):
            self.row(r).write_blanks(c1, c2,  style)
        self.__merged_ranges.append((r1, r2, c1, c2))

    def insert_bitmap(self, filename, row, col, x = 0, y = 0, scale_x = 1, scale_y = 1):
        bmp = Bitmap.ImDataBmpRecord(filename)
        obj = Bitmap.ObjBmpRecord(row, col, self, bmp, x, y, scale_x, scale_y)

        self.__bmp_rec += obj.get() + bmp.get()

    def insert_picture(self, img_path, img_type, img_data, w, h, row=0, col=0, x = 0, y = 0, scale_x = 1, scale_y = 1):
        self._parent.drawing_group.insert(self.__idx, img_type, img_data)
        cnt = self._parent.drawing_group.get_count()
        #pic_cnt = self._parent.drawing_group.get_pic_count_in_sheet(self.__idx)
        self.__pic_rec.insert(cnt, img_type, w, h, row, col, x, y, scale_x, scale_y)

    def col(self, indx):
        if indx not in self.__cols:
            self.__cols[indx] = Column(indx, self)
        return self.__cols[indx]

    def row(self, indx):
        if indx not in self.__rows:
            if indx in self.__flushed_rows:
                raise Exception("Attempt to reuse row index %d of sheet %r after flushing" % (indx, self.__name))
            self.__rows[indx] = self.Row(indx, self)
            if indx > self.last_used_row:
                self.last_used_row = indx
            if indx < self.first_used_row:
                self.first_used_row = indx
        return self.__rows[indx]

    def set_row_height(self, rowno, height): # in pixels
        row = self.__rows.get(rowno)
        if row:
            row._height_in_pixels = height
            twips = int((float(height) - 2/5) * 50/83 * 20.0)
            row.height = twips
            row.height_mismatch = 1

    def row_height(self, row): # in pixels
        if row in self.__rows:
            return self.__rows[row]._height_in_pixels
        else:
            return 17

    def col_width(self, col): # in pixels
        if col in self.__cols:
            return self.__cols[col].width_in_pixels()
        else:
            return 64


    ##################################################################
    ## BIFF records generation
    ##################################################################

    def __bof_rec(self):
        return BIFFRecords.Biff8BOFRecord(BIFFRecords.Biff8BOFRecord.WORKSHEET).get()

    def __update_row_visible_levels(self):
        if self.__rows:
            temp = max([self.__rows[r].level for r in self.__rows]) + 1
            self.__row_visible_levels = max(temp, self.__row_visible_levels)

    def __guts_rec(self):
        self.__update_row_visible_levels()
        col_visible_levels = 0
        if len(self.__cols) != 0:
            col_visible_levels = max([self.__cols[c].level for c in self.__cols]) + 1
        return BIFFRecords.GutsRecord(
            self.__row_gut_width, self.__col_gut_height, self.__row_visible_levels, col_visible_levels).get()

    def __defaultrowheight_rec(self):
        options = 0x0000
        options |= (self.row_default_height_mismatch & 1) << 0
        options |= (self.row_default_hidden & 1) << 1
        options |= (self.row_default_space_above & 1) << 2
        options |= (self.row_default_space_below & 1) << 3
        defht = self.__row_default_height
        return BIFFRecords.DefaultRowHeightRecord(options, defht).get()

    def __wsbool_rec(self):
        options = 0x00
        options |= (self.__show_auto_page_breaks & 0x01) << 0
        options |= (self.__dialogue_sheet & 0x01) << 4
        options |= (self.__auto_style_outline & 0x01) << 5
        options |= (self.__outline_below & 0x01) << 6
        options |= (self.__outline_right & 0x01) << 7
        options |= (self.__fit_num_pages & 0x01) << 8
        options |= (self.__show_row_outline & 0x01) << 10
        options |= (self.__show_col_outline & 0x01) << 11
        options |= (self.__alt_expr_eval & 0x01) << 14
        options |= (self.__alt_formula_entries & 0x01) << 15

        return BIFFRecords.WSBoolRecord(options).get()

    def __eof_rec(self):
        return BIFFRecords.EOFRecord().get()

    def __colinfo_rec(self):
        result = ''
        for col in self.__cols:
            result += self.__cols[col].get_biff_record()
        return result

    def __dimensions_rec(self):
        return BIFFRecords.DimensionsRecord(
            self.first_used_row, self.last_used_row,
            self.first_used_col, self.last_used_col
            ).get()

    def __window2_rec(self):
        # Appends SCL record.
        options = 0
        options |= (self.__show_formulas        & 0x01) << 0
        options |= (self.__show_grid            & 0x01) << 1
        options |= (self.__show_headers         & 0x01) << 2
        options |= (self.__panes_frozen         & 0x01) << 3
        options |= (self.show_zero_values       & 0x01) << 4
        options |= (self.__auto_colour_grid     & 0x01) << 5
        options |= (self.__cols_right_to_left   & 0x01) << 6
        options |= (self.__show_outline         & 0x01) << 7
        options |= (self.__remove_splits        & 0x01) << 8
        options |= (self.__selected             & 0x01) << 9
        options |= (self.__sheet_visible        & 0x01) << 10
        options |= (self.__page_preview         & 0x01) << 11
        if self.explicit_magn_setting:
            # Experimentation: caller can set the scl magn.
            # None -> no SCL record written
            # Otherwise 10 <= scl_magn <= 400 or scl_magn == 0
            # Note: value 0 means use 100 for normal view, 60 for page break preview
            # BREAKING NEWS: Excel interprets scl_magn = 0 very literally, your
            # sheet appears like a tiny dot on the screen
            scl_magn = self.__scl_magn
        else:
            if self.__page_preview:
                scl_magn = self.__preview_magn
                magn_default = 60
            else:
                scl_magn = self.__normal_magn
                magn_default = 100
            if scl_magn == magn_default or scl_magn == 0:
                # Emulate what we think MS does
                scl_magn = None # don't write an SCL record
        return BIFFRecords.Window2Record(
            options, self.__first_visible_row, self.__first_visible_col,
            self.__grid_colour,
            self.__preview_magn, self.__normal_magn, scl_magn).get()

    def __panes_rec(self):
        if self.__vert_split_pos is None and self.__horz_split_pos is None:
            return ""

        if self.__vert_split_pos is None:
            self.__vert_split_pos = 0
        if self.__horz_split_pos is None:
            self.__horz_split_pos = 0

        if self.__panes_frozen:
            if self.__vert_split_first_visible is None:
                self.__vert_split_first_visible = self.__vert_split_pos
            if self.__horz_split_first_visible is None:
                self.__horz_split_first_visible = self.__horz_split_pos
        else:
            if self.__vert_split_first_visible is None:
                self.__vert_split_first_visible = 0
            if self.__horz_split_first_visible is None:
                self.__horz_split_first_visible = 0
            if not self.split_position_units_are_twips:
                # inspired by pyXLWriter
                if self.__horz_split_pos > 0:
                    self.__horz_split_pos = 20 * self.__horz_split_pos + 255
                if self.__vert_split_pos > 0:
                    self.__vert_split_pos = 113.879 * self.__vert_split_pos + 390

        result = BIFFRecords.PanesRecord(*map(int, (
            self.__vert_split_pos,
            self.__horz_split_pos,
            self.__horz_split_first_visible,
            self.__vert_split_first_visible,
            self.active_pane
            ))).get()

        return result

    def __row_blocks_rec(self):
        result = []
        for row in self.__rows.itervalues():
            result.append(row.get_row_biff_data())
            result.append(row.get_cells_biff_data())
        return ''.join(result)

    def __merged_rec(self):
        return BIFFRecords.MergedCellsRecord(self.__merged_ranges).get()

    def __bitmaps_rec(self):
        return self.__bmp_rec

    def __picture_rec(self):
        return self.__pic_rec.get()

    def __calc_settings_rec(self):
        result = ''
        result += BIFFRecords.CalcModeRecord(self.__calc_mode & 0x01).get()
        result += BIFFRecords.CalcCountRecord(self.__calc_count & 0xFFFF).get()
        result += BIFFRecords.RefModeRecord(self.__RC_ref_mode & 0x01).get()
        result += BIFFRecords.IterationRecord(self.__iterations_on & 0x01).get()
        result += BIFFRecords.DeltaRecord(self.__delta).get()
        result += BIFFRecords.SaveRecalcRecord(self.__save_recalc & 0x01).get()
        return result

    def __print_settings_rec(self):
        result = ''
        result += BIFFRecords.PrintHeadersRecord(self.__print_headers).get()
        result += BIFFRecords.PrintGridLinesRecord(self.__print_grid).get()
        result += BIFFRecords.GridSetRecord(self.__grid_set).get()
        result += BIFFRecords.HorizontalPageBreaksRecord(self.__horz_page_breaks).get()
        result += BIFFRecords.VerticalPageBreaksRecord(self.__vert_page_breaks).get()
        result += BIFFRecords.HeaderRecord(self.__header_str).get()
        result += BIFFRecords.FooterRecord(self.__footer_str).get()
        result += BIFFRecords.HCenterRecord(self.__print_centered_horz).get()
        result += BIFFRecords.VCenterRecord(self.__print_centered_vert).get()
        result += BIFFRecords.LeftMarginRecord(self.__left_margin).get()
        result += BIFFRecords.RightMarginRecord(self.__right_margin).get()
        result += BIFFRecords.TopMarginRecord(self.__top_margin).get()
        result += BIFFRecords.BottomMarginRecord(self.__bottom_margin).get()

        setup_page_options =  (self.__print_in_rows & 0x01) << 0
        setup_page_options |=  (self.__portrait & 0x01) << 1
        setup_page_options |=  (0x00 & 0x01) << 2
        setup_page_options |=  (self.__print_not_colour & 0x01) << 3
        setup_page_options |=  (self.__print_draft & 0x01) << 4
        setup_page_options |=  (self.__print_notes & 0x01) << 5
        setup_page_options |=  (0x00 & 0x01) << 6
        setup_page_options |=  (0x01 & 0x01) << 7
        setup_page_options |=  (self.__print_notes_at_end & 0x01) << 9
        setup_page_options |=  (self.__print_omit_errors & 0x03) << 10

        result += BIFFRecords.SetupPageRecord(self.__paper_size_code,
                                self.__print_scaling,
                                self.__start_page_number,
                                self.__fit_width_to_pages,
                                self.__fit_height_to_pages,
                                setup_page_options,
                                self.__print_hres,
                                self.__print_vres,
                                self.__header_margin,
                                self.__footer_margin,
                                self.__copies_num).get()
        return result

    def __protection_rec(self):
        result = ''
        result += BIFFRecords.ProtectRecord(self.__protect).get()
        result += BIFFRecords.ScenProtectRecord(self.__scen_protect).get()
        result += BIFFRecords.WindowProtectRecord(self.__wnd_protect).get()
        result += BIFFRecords.ObjectProtectRecord(self.__obj_protect).get()
        result += BIFFRecords.PasswordRecord(self.__password).get()
        return result

    def get_biff_data(self):
        result = [
            self.__bof_rec(),
            self.__calc_settings_rec(),
            self.__guts_rec(),
            self.__defaultrowheight_rec(),
            self.__wsbool_rec(),
            self.__colinfo_rec(),
            self.__dimensions_rec(),
            self.__print_settings_rec(),
            self.__protection_rec(),
            ]
        if self.row_tempfile:
            with open(self.row_tempfile, 'rb') as f:
                #self.row_tempfile.flush()
                #self.row_tempfile.seek(0)
                result.append(f.read())
        result.extend([
            self.__row_blocks_rec(),
            self.__merged_rec(),
            self.__bitmaps_rec(),
            self.__picture_rec(),
            self.__window2_rec(),
            self.__panes_rec(),
            self.__eof_rec(),
            ])
        return ''.join(result)

    def flush_row_data(self):
        if self.row_tempfile is None:
            #f = tempfile.NamedTemporaryFile()
            fd, self.row_tempfile = tempfile.mkstemp(suffix='.sheet')
            with os.fdopen(fd, 'wb') as f:
                f.write(self.__row_blocks_rec())
        else:
            with open(self.row_tempfile, 'wb') as f:
                f.write(self.__row_blocks_rec())
        for rowx in self.__rows:
            self.__flushed_rows[rowx] = 1
        self.__update_row_visible_levels()
        self.__rows = {}

    def insert_row_before(self, int idx, int nrows=1):
        cdef Row row
        cdef Cell cell
        cdef int i, rowidx, new_idx
        cdef dict _rows = self.__rows

        indexes = (x for x in _rows if x >= idx)
        for rowidx in reversed(sorted(indexes)):
            row = self.__rows[rowidx]
            new_idx = rowidx + nrows
            row._idx = new_idx
            for cell in row._cells.values():
                if cell is not None:
                    try:
                        cell.rowx = new_idx 
                    except:
                        pass
            self.__rows[new_idx] = row
            self.__rows[rowidx] = None

        for i in range(idx, idx + nrows):
            row = Row(i, self)
            self.__rows[i] = row

        self.update_ranges(idx, nrows)

    def update_ranges(self, int idx, int nrows=1):
        for i, r in enumerate(self.__merged_ranges):
            r1, r2, c1, c2 = r
            if r1 >= idx:
                self.__merged_ranges[i] = (r1 + nrows, r2 + nrows, c1, c2)
            elif r2 >= idx:
                self.__merged_ranges[i] = (r1, r2 + nrows, c1, c2)

    def set_print_title(self, r1, r2):
        self._parent.add_print_title(self.__idx, r1, r2)

    cpdef get_copy(self, name, parent=None):
        cdef Worksheet sht
        cdef Row row
        cdef int indx

        if parent is None:
            parent = self._parent
        sht = Worksheet(name, parent, cell_overwrite_ok=True)
        for indx, row in self.__rows.iteritems():
            sht.__rows[indx] = row.get_copy(indx, sht)

        for indx, col in self.__cols.iteritems():
            sht.__cols[indx] = col.get_copy(indx, sht)

        sht.__merged_ranges = self.__merged_ranges[:]
        sht.__bmp_rec = self.__bmp_rec
        #for pic in self.__pic_rec.pics:
        #    sht.__pic_rec.insert_pic(*pic)

        #print '??'
        sht.__show_formulas = self.__show_formulas
        sht.__show_grid = self.__show_grid
        sht.__show_headers = self.__show_headers
        sht.__panes_frozen = self.__panes_frozen
        sht.show_zero_values = self.show_zero_values
        sht.__auto_colour_grid = self.__auto_colour_grid
        sht.__cols_right_to_left = self.__cols_right_to_left
        sht.__show_outline = self.__show_outline
        sht.__remove_splits = self.__remove_splits
        sht.__selected = self.__selected
        sht.__sheet_visible = self.__sheet_visible
        sht.__page_preview = self.__page_preview

        sht.__first_visible_row = self.__first_visible_row
        sht.__first_visible_col = self.__first_visible_col
        sht.__grid_colour = self.__grid_colour
        sht.__preview_magn = self.__preview_magn
        sht.__normal_magn = self.__normal_magn
        sht.__scl_magn = self.__scl_magn
        sht.explicit_magn_setting = self.explicit_magn_setting

        sht.visibility = self.visibility

        sht.__vert_split_pos = self.__vert_split_pos
        sht.__horz_split_pos = self.__horz_split_pos
        sht.__vert_split_first_visible = self.__vert_split_first_visible
        sht.__horz_split_first_visible = self.__horz_split_first_visible

        sht.split_position_units_are_twips = self.split_position_units_are_twips

        sht.__row_gut_width = self.__row_gut_width
        sht.__col_gut_height = self.__col_gut_height

        sht.__show_auto_page_breaks = self.__show_auto_page_breaks
        sht.__dialogue_sheet = self.__dialogue_sheet
        sht.__auto_style_outline = self.__auto_style_outline
        sht.__outline_below = self.__outline_below
        sht.__outline_right = self.__outline_right
        sht.__fit_num_pages = self.__fit_num_pages
        sht.__show_row_outline = self.__show_row_outline
        sht.__show_col_outline = self.__show_col_outline
        sht.__alt_expr_eval = self.__alt_expr_eval
        sht.__alt_formula_entries = self.__alt_formula_entries

        sht.__row_default_height = self.__row_default_height
        sht.row_default_height_mismatch = self.row_default_height_mismatch
        sht.row_default_hidden = self.row_default_hidden
        sht.row_default_space_above = self.row_default_space_above
        sht.row_default_space_below = self.row_default_space_below

        sht.__col_default_width = self.__col_default_width

        sht.__calc_mode = self.__calc_mode
        sht.__calc_count = self.__calc_count
        sht.__RC_ref_mode = self.__RC_ref_mode
        sht.__iterations_on = self.__iterations_on
        sht.__delta = self.__delta
        sht.__save_recalc = self.__save_recalc

        sht.__print_headers = self.__print_headers
        sht.__print_grid = self.__print_grid
        sht.__grid_set = self.__grid_set
        sht.__vert_page_breaks = self.__vert_page_breaks[:]
        sht.__horz_page_breaks = self.__horz_page_breaks[:]
        sht.__header_str = self.__header_str
        sht.__footer_str = self.__footer_str
        sht.__print_centered_vert = self.__print_centered_vert
        sht.__print_centered_horz = self.__print_centered_horz
        sht.__left_margin = self.__left_margin
        sht.__right_margin = self.__right_margin
        sht.__top_margin = self.__top_margin
        sht.__bottom_margin = self.__bottom_margin
        sht.__paper_size_code = self.__paper_size_code
        sht.__print_scaling = self.__print_scaling
        sht.__start_page_number = self.__start_page_number
        sht.__fit_width_to_pages = self.__fit_width_to_pages
        sht.__fit_height_to_pages = self.__fit_height_to_pages
        sht.__print_in_rows = self.__print_in_rows
        sht.__portrait = self.__portrait
        sht.__print_not_colour = self.__print_not_colour
        sht.__print_draft = self.__print_draft
        sht.__print_notes = self.__print_notes
        sht.__print_notes_at_end = self.__print_notes_at_end
        sht.__print_omit_errors = self.__print_omit_errors
        sht.__print_hres = self.__print_hres
        sht.__print_vres = self.__print_vres
        sht.__header_margin = self.__header_margin
        sht.__footer_margin = self.__footer_margin
        sht.__copies_num = self.__copies_num

        sht.__wnd_protect = self.__wnd_protect
        sht.__obj_protect = self.__obj_protect
        sht.__protect = self.__protect
        sht.__scen_protect = self.__scen_protect
        sht.__password = self.__password

        sht.last_used_row = self.last_used_row
        sht.first_used_row = self.first_used_row
        sht.last_used_col = self.last_used_col
        sht.first_used_col = self.first_used_col
        sht.row_tempfile = self.row_tempfile
        sht.__flushed_rows = self.__flushed_rows.copy()
        sht.__row_visible_levels = self.__row_visible_levels

        return sht
 
