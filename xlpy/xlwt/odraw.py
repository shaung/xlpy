# -*- coding: utf-8 -*-

from BIFFRecords import BiffRecord
from struct import *
import Bitmap

class ODrawRecordBase:
    def get_full_len(self):
        len = self.get_len()
        return len > 0 and (len + 8) or 0

    def get_len(self):
        if self._len < 0:
            return self._get_len()
        return self._len

    def get_header(self):
        rslt = ''
        ver = self._ver | self._ins << 4
        rslt = pack('<HHL', ver, self._type, self.get_len())
        return rslt

class MSODrawingGroupRecord(BiffRecord):
    _REC_ID = 0x00EB

    def __init__(self):
        self.dgct = OfficeArtDggContainer()

    def add_sheet(self, sheet_id, pic_count):
        self.dgct.add_sheet(sheet_id, pic_count)

    def insert(self, sheet_id, pic_type, img_data):
        self.dgct.insert(sheet_id, pic_type, img_data)
        self.finish()

    def get_count(self):
        return self.dgct.get_count()

    def get_pic_count_in_sheet(self, sheet_id):
        return self.dgct.drawing_group.get_pic_count_in_sheet(sheet_id)
 
    def finish(self):
        self._rec_data = self.dgct.get()

class PictureSection:
    def __init__(self, sheet_id, sheet):
        self.sheet_id = sheet_id
        self.sheet = sheet
        self.objs = []
        first = MSODrawingRecordFirst(sheet_id, None, 0, None)
        obj = UndefinedBase()
        self.objs.append((first, obj))

        self.pics = []

        self.is_empty = True

    def insert(self, pic_id, pic_type, width, height, row, col, x=0, y=0, scale_x=1, scale_y=1):
        self.pics.append((pic_id, pic_type, width, height, row, col, x, y, scale_x, scale_y))

        shape_id = len(self.pics) + 1

        # Scale the frame of the image.
        width = width * scale_x
        height = height * scale_y

        # Calculate the vertices of the image and write the OBJ record
        coordinates = Bitmap._position_image(self.sheet, row, col, x, y, width, height)
        #print coordinates

        if self.is_empty:
            self.objs = []
            odraw_rec = MSODrawingRecordFirst(self.sheet_id, pic_id, shape_id, pic_type, *coordinates)
            self.is_empty = False
        else:
            odraw_rec = MSODrawingRecordFollow(self.sheet_id, pic_id, shape_id, pic_type, *coordinates)
        obj = ObjGraphicRecord(pic_id)
        self.objs.append([odraw_rec, obj])
        first_rec = self.objs[0][0]
        first_rec.add()

    def finish(self):
        for x, o in self.objs:
            x.finish()

    def get(self):
        if self.is_empty:
            return ''
        self.finish()
        rslt = ''
        rslt += ''.join([x.get() + o.get() for (x, o) in self.objs])
        return rslt

class MSODrawingRecordBase(BiffRecord):
    _REC_ID = 0x00EC

class MSODrawingRecordFirst(MSODrawingRecordBase):
    def __init__(self, sheet_id, pic_id, shape_id, pic_type, *coordinates):
        self.dgct = OfficeArtDgContainer(sheet_id)
        if pic_id is not None:
            self.dgct.insert(pic_id, shape_id, pic_type, *coordinates)

    def add(self):
        self.dgct.add()

    def finish(self):
        self._rec_data = self.dgct.get()

class MSODrawingRecordFollow(MSODrawingRecordBase):
    def __init__(self, sheet_id, pic_id, shape_id, pic_type, *coordinates):
        self.spct = OfficeArtSpContainer(sheet_id)
        self.spct.insert(pic_id, shape_id, pic_type, *coordinates)
        self.finish()

    def finish(self):
        self._rec_data = self.spct.get()

class ObjGraphicRecord(BiffRecord):
    _REC_ID = 0x005D    # Record identifier

    def __init__(self, pic_id):
        ft = 0x0015        # common object data
        ot = 0x0008        # Object type. 8 = Picture
        id = 0x0000 | pic_id
        grbit = 0x6011
        cb = 0x0012

        data = ''
        data += pack('<H', ft)
        data += pack('<H', cb)
        data += pack('<H', ot)
        data += pack('<H', id)
        data += pack('<H', grbit)
        data += pack('<3L', 0, 0, 0)
        data += pack('<L', 0)
        self._rec_data = data

"""
    The entrance
"""
class OfficeArtDggContainer(ODrawRecordBase):
    _ver = 0xF
    _ins = 0x0
    _type = 0xF000
    _len = -1

    def __init__(self):
        # ODrawRecordBase.__init__(self)
        # drawingGroup (variable): An OfficeArtFDGGBlock record, as defined in section 2.2.48, that
        # specifies document-wide information about all the drawings that are saved in the file.
        self.drawing_group = OfficeArtFDGGBlock()
        # blipStore (variable): An OfficeArtBStoreContainer record, as defined in section 2.2.20, that
        # specifies the container for all the BLIPs that are used in all the drawings in the parent
        # document.
        self.blip_store = OfficeArtBStoreContainer()
        # drawingPrimaryOptions (variable): An OfficeArtFOPT record, as defined in section 2.2.9,
        # that specifies the default properties for all drawing objects that are contained in all the
        # drawings in the parent document.
        self.opt1 = OfficeArtFOPT()
        # drawingTertiaryOptions (variable): An OfficeArtTertiaryFOPT record, as defined in section
        # 2.2.11, that specifies the default properties for all the drawing objects that are contained in all
        # the drawings in the parent document.
        # colorMRU (variable): An OfficeArtColorMRUContainer record, as defined in section 2.2.43,
        # that specifies the most recently used custom colors.
        # splitColors (variable): An OfficeArtSplitMenuColorContainer record, as defined in section
        # 2.2.45, that specifies a container for the colors that were most recently used to format
        # shapes.
        self.split_colors = OfficeArtSplitMenuColorContainer()

        self.init_props()

    def init_props(self):
        self.opt1.insert(0x00BF, 0x00080008)
        self.opt1.insert(0x0181, 0x08000009)
        self.opt1.insert(0x01C0, 0x08000040)

    def get_count(self):
        return self.drawing_group.get_count()

    def _get_len(self):
        rslt = 0
        rslt += self.drawing_group.get_full_len()
        rslt += self.blip_store.get_full_len()
        rslt += self.opt1.get_full_len()
        rslt += self.split_colors.get_full_len()
        return rslt

    def get(self):
        if self.get_count() == 0:
            return ''
        rslt = self.get_header()
        if self.drawing_group:
            rslt += self.drawing_group.get()
        if self.blip_store:
            rslt += self.blip_store.get()
        rslt += self.opt1.get()
        rslt += self.split_colors.get()
        return rslt

    def add_sheet(self, sheet_id, pic_count):
        self.drawing_group.add_sheet(sheet_id, pic_count)

    def insert(self, sheet_id, pic_type, img_data):
        self.drawing_group.insert(sheet_id)
        self.blip_store.insert(pic_type, img_data)

class OfficeArtFDGGBlock(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x0
    _type = 0xF006
    _len = -1

    _count = 0

    def __init__(self):
        # ODrawRecordBase.__init__(self)
        self.head = OfficeArtFDGG()
        #self.rgidcl = []
        self.idcl_map = {} # maintain the sheets id
        #self.first_idcl = OfficeArtIDCL(0x01, 0x01)
        #self.rgidcl.append(self.first_idcl)
        #self.rgidcl.append(OfficeArtIDCL(0x00, 0x04))

    def add_sheet(self, sheet_id, pic_count):
        idcl = OfficeArtIDCL(sheet_id, pic_count + 1)
        self.idcl_map[sheet_id] = idcl
        if sheet_id == max(self.idcl_map):
            self.head.spidmax = (sheet_id + 1) * 4 << 0x08 | idcl.cspid_cur

    def insert(self, sheet_id, count=1):
        if sheet_id not in self.idcl_map:
            self.idcl_map[sheet_id] = OfficeArtIDCL(sheet_id, 1)
        idcl = self.idcl_map[sheet_id]
        idcl.cspid_cur += 1
        if sheet_id == max(self.idcl_map):
            self.head.spidmax = (sheet_id + 1) * 4 << 0x08 | idcl.cspid_cur
        self._count += 1
        self.head.cidcl = len(self.idcl_map) + 1
        self.head.cspsaved += 0x1
        self.head.cdgsaved = len(self.idcl_map)
        #if self._count > 1:
        #self.first_idcl.cspid_cur += 0x1

    def get_cidcl(self):
        return self.head.cidcl

    def get_count(self):
        return sum([v.cspid_cur - 1 for v in self.idcl_map.values()])

    def get_pic_count_in_sheet(self, sheet_id):
        return self.idcl_map[sheet_id].cspid_cur

    def _get_len(self):
        return 0x10 + (self.get_cidcl() - 1) * 0x08

    def get(self):
        rslt = self.get_header()
        rslt += self.head.get()
        sheets = [k for k in self.idcl_map]
        sheets.sort()
        rslt += ''.join([self.idcl_map[x].get() for x in sheets])
        return rslt

class OfficeArtFDGG:
    def __init__(self, count=0):
        # spidMax (4 bytes): An MSOSPID structure, as defined in section 2.1.2, specifying the current
        # maximum shape identifier that is used in any drawing. This value MUST be less than
        # 0x03FFD7FF.
        # (sheet idx + 1) * 4 << 8 + max_image count
        # high 24-bit stands for max image counts in a sheet
        self.spidmax = 0x0000EC01 # should be safe
        # cidcl (4 bytes): An unsigned integer that specifies the number of OfficeArtIDCL records, as
        # defined in section 2.2.46, + 1. This value MUST be less than 0x0FFFFFFF.
        self.cidcl = 0x02 # sheet count + 1
        # cspSaved (4 bytes): An unsigned integer specifying the total number of shapes that have been
        # saved in all of the drawings.
        self.cspsaved = 0x01 + count # sheet count + count of images
        # cdgSaved (4 bytes): An unsigned integer specifying the total number of drawings that have
        # been saved in the file.
        self.cdgsaved = 0x01 # count of sheets

    def get(self):
        rslt = ''
        rslt += pack('<L', self.spidmax)
        rslt += pack('<L', self.cidcl)
        rslt += pack('<L', self.cspsaved)
        rslt += pack('<L', self.cdgsaved)
        return rslt

class OfficeArtIDCL:
    def __init__(self, dgid=0x00, cspid_cur=0x00):
        # dgid (4 bytes): An MSODGID structure, as defined in section 2.1.1, specifying the drawing
        # identifier that owns this identifier cluster.
        self.dgid = dgid
        # cspidCur (4 bytes): An unsigned integer that, if less than 0x00000400, specifies the largest
        # shape identifier that is currently assigned in this cluster, or that otherwise specifies that no
        # shapes can be added to the drawing.
        self.cspid_cur = cspid_cur

    def get(self):
        rslt = ''
        rslt += pack('<L', self.dgid)
        rslt += pack('<L', self.cspid_cur)
        return rslt

class OfficeArtBStoreContainer(ODrawRecordBase):
    _ver = 0xF
    _ins = 0x0
    _type = 0xF001
    _len = -1

    def __init__(self):
        # ODrawRecordBase.__init__(self)
        self.rgfbs = []  # per file

    def _get_len(self):
        return sum([x.get_full_len() for x in self.rgfbs])

    def get(self):
        rslt = self.get_header()
        rslt += ''.join([x.get() for x in self.rgfbs])
        return rslt

    def insert(self, pic_type, img_data):
        self._ins += 0x1
        fbse = OfficeArtFBSE()
        fbse.insert(pic_type, img_data)
        self.rgfbs.append(fbse)

class OfficeArtBStoreContainerFileBlock(ODrawRecordBase):
    pass

class MSOBLIPTYPE:
    """
    MSOBLIPTYPE enumeration, as shown in the following table, specifies the persistence format
    of bitmap data.
    """
    msoblipERROR = 0x00 # Error reading the file.
    msoblipUNKNOWN = 0x01 # Unknown BLIP type.
    msoblipEMF = 0x02 #EMF.
    msoblipWMF = 0x03 # WMF.
    msoblipPICT = 0x04 # Macintosh PICT.
    msoblipJPEG = 0x05 # JPEG.
    msoblipPNG = 0x06 # PNG.
    msoblipDIB = 0x07 # DIB
    msoblipTIFF = 0x11 # TIFF
    msoblipCMYKJPEG = 0x12 # JPEG in the YCCK or CMYK color space.

class OfficeArtFBSE(ODrawRecordBase):
    _ver = 0x2
    _ins = 0x6
    _type = 0xF007
    _len = -1

    def __init__(self):
        # ODrawRecordBase.__init__(self)
        # 1b
        self.bt_win32 = MSOBLIPTYPE.msoblipPNG # 06
        # 1b
        self.bt_macos = MSOBLIPTYPE.msoblipPNG # 06
        # 16b An MD4 message digest, as specified in [RFC1320], that specifies the unique identifier of the pixel data in the BLIP.
        self.rgbuid = 0x0 # should use blip's uid1
        # 2b An unsigned integer that specifies an application-defined internal resource tag.  This value MUST be 0xFF for external files.
        self.tag = 0x00 # fixed
        # 4b An unsigned integer that specifies the size, in bytes, of the BLIP in the stream.
        self.size = 0x00 # should use blip's size
        # 4b An unsigned integer that specifies the number of references to the BLIP. A value of 0x00000000 specifies an empty slot in the OfficeArtBStoreContainer record, as defined in section 2.2.20.
        self.cref = 0x01 # seems to be fixed
        # 4b An MSOFO structure, as defined in section 2.1.4, that specifies the file
        #    offset into the associated OfficeArtBStoreDelay record, as defined in section 2.2.21, (delay
        #    stream). A value of 0xFFFFFFFF specifies that the file is not in the delay stream, and in this
        #    case, cRef MUST be 0x00000000.
        self.fo_delay = 0x0 # fixed
        # 1b
        self.unused1 = 0x0
        # 1b An unsigned integer that specifies the length, in bytes, of the nameData
        #    field, including the terminating NULL character. This value MUST be an even number and less
        #    than or equal to 0xFE. If the value is 0x00, nameData will not be written.
        self.cb_name = 0x0 # fixed to 0 since we do not use the name_data
        # 1b
        self.unused2 = 0x00
        # 1b
        self.unused3 = 0x00
        # var A Unicode null-terminated string that specifies the name of the BLIP.
        self.name_data = '' # fixed empty
        # var An OfficeArtBlip record, as defined in section 2.2.23, specifying the
        #     BLIP file data that is embedded in this record. If this value is not 0, foDelay MUST be ignored.
        self.embedded_blip = None

    def insert_png(self, img_data):
        self._ins = MSOBLIPTYPE.msoblipPNG
        self.bt_win32 = MSOBLIPTYPE.msoblipPNG
        self.bt_macos = MSOBLIPTYPE.msoblipPNG
        self.embedded_blip = OfficeArtBlipPNG()
        #self.embedded_blip.load_from_file(fpath)
        self.embedded_blip.load_from_string(img_data)

    def insert_jpeg(self, img_data):
        self._ins = MSOBLIPTYPE.msoblipJPEG
        self.bt_win32 = MSOBLIPTYPE.msoblipJPEG
        self.bt_macos = MSOBLIPTYPE.msoblipJPEG
        self.embedded_blip = OfficeArtBlipJPEG()
        #self.embedded_blip.load_from_file(fpath)
        self.embedded_blip.load_from_string(img_data)

    def insert(self, pic_type, *args):
        mapper = {
            'png': self.insert_png,
            'jpeg': self.insert_jpeg,
            'jpg': self.insert_jpeg,
        }
        func = mapper[pic_type]
        func(*args)

    def get_size(self):
        return self.embedded_blip.get_full_len()

    def _get_len(self):
        return len(self.name_data) + self.get_size() + 36

    def get(self):
        rslt = self.get_header()
        rslt += pack('<B', self.bt_win32)
        rslt += pack('<B', self.bt_macos)
        rslt += pack('<16s', self.embedded_blip.rgb_uid1)
        rslt += pack('<H', self.tag)
        rslt += pack('<L', self.get_size())
        rslt += pack('<L', self.cref)
        rslt += pack('<L', self.fo_delay)
        rslt += pack('<B', self.unused1)
        rslt += pack('<B', self.cb_name)
        rslt += pack('<B', self.unused2)
        rslt += pack('<B', self.unused3)
        rslt += self.name_data
        rslt += self.embedded_blip.get()
        return rslt

class OfficeArtBlip(ODrawRecordBase):
    def __init__(self):
        # OfficeArtBlip.__init__(self)
        # rgbUid1 (16 bytes): An MD4 message digest, as specified in [RFC1320], that specifies the
        # unique identifier of the uncompressed BLIPFileData.
        self.rgb_uid1 = 0x00
        # rgbUid2 (16 bytes): An MD4 message digest, as specified in [RFC1320], that specifies the
        # unique identifier of the uncompressed BLIPFileData. This field only exists if recInstance
        # equals 0x6E1. If this value exists, rgbUid1 MUST be ignored.
        self.rgb_uid2 = 0x00
        # tag (1 byte): An unsigned integer that specifies an application-defined internal resource tag.
        # This value MUST be 0xFF for external files.
        self.tag = 0xFF
        self.data = ''

    def _get_len(self):
        return len(self.data) + self.get_ins_len()

    def load_from_file(self, fpath):
        f = open(fpath, 'rb')
        self.data = f.read()
        f.close()
        self.rgb_uid1 = self.make_md4()

    def load_from_string(self, img_data):
        self.data = img_data
        self.rgb_uid1 = self.make_md4()

    def make_md4(self):
        import hashlib
        import binascii
        return binascii.unhexlify(hashlib.new('md4', self.data).hexdigest().upper())
        #from Crypto.Hash import MD4
        #return MD4.new(self.data).digest()

    def get(self):
        rslt = self.get_header()
        rslt += pack('<16s', self.rgb_uid1)
        if self.rgb_uid2:
            rslt += pack('<16s',self.rgb_uid2)
        rslt += pack('<B', self.tag)
        rslt += self.data
        return rslt

class OfficeArtBlipPNG(OfficeArtBlip):
    _ver = 0x0
    _ins = 0x6E0 # A value of 0x6E0 to specify one UID, or a value of 0x6E1 to specify two UIDs.
    _type = 0xF01E
    _len = -1

    def get_ins_len(self):
        return self._ins == 0x6E0 and 17 or 33

class OfficeArtBlipJPEG(OfficeArtBlip):
    _ver = 0x0
    _ins = 0x6E0 # A value of 0x6E0 to specify one UID, or a value of 0x6E1 to specify two UIDs.
    _type = 0xF01D
    _len = -1

    _color_mode = 'rgb1'
    COLOR_MODES = {
        'rgb1': 0x46A,
        'rgb2': 0x46B,
        'cmyk1': 0x6E2,
        'cmyk2': 0x6E3,
    }

    def __init__(self, color_mode='rgb1'):
        assert(color_mode in self.COLOR_MODES)
        self._color_mode = color_mode
        self._ins = self.COLOR_MODES[color_mode]
        OfficeArtBlip.__init__(self)

    def get_ins_len(self):
        return self._color_mode in ('rgb1', 'cmyk1') and 17 or 33

class OfficeArtOPTBase(ODrawRecordBase):
    def __init__(self):
        self.fopt = OfficeArtRGFOPTE()

    def insert(self, k, v, fbid=0, fcomplex=0):
        self._ins += 1
        self.fopt.insert(k, v, fbid, fcomplex)

    def _get_len(self):
        return self.fopt.get_len()

    def get(self):
        if self.fopt.get_size() == 0:
            return ''
        rslt = self.get_header()
        rslt += self.fopt.get()
        return rslt

class OfficeArtFOPT(OfficeArtOPTBase):
    _ver = 0x3
    _ins = 0x0 # count of properties
    _type = 0xF00B
    _len = -1

class OfficeArtSecondaryFOPT(OfficeArtOPTBase):
    _ver = 0x3
    _ins = 0x0 # count of properties
    _type = 0xF121
    _len = -1

class OfficeArtTertiaryFOPT(OfficeArtOPTBase):
    _ver = 0x3
    _ins = 0x0 # count of properties
    _type = 0xF122
    _len = -1

class OfficeArtRGFOPTE:
    def __init__(self):
        self.props = []
        self.datas = []

    def insert(self, k, v, fbid, fcomplex):
        prop = OfficeArtFOPTE(k, v, fbid, fcomplex)
        self.props.append(prop)
        if fcomplex:
            self.datas.append(v)

    def get_size(self):
        return len(self.props)

    def get_len(self):
        return sum([x.get_len() for x in self.props])
        # + sum([x.get_len() for x in self.datas])

    def get(self):
        rslt = ''
        rslt += ''.join([x.get() for x in self.props])
        rslt += ''.join(self.datas)
        return rslt

class OfficeArtFOPTE:
    def __init__(self, k, op, fbid, fcomplex):
        # opid (2 bytes): An OfficeArtFOPTEOPID record, as defined in section 2.2.8, that specifies the
        # header information for this property.
        self.opid = OfficeArtFOPTEOPID(k, fbid, fcomplex)
        # op (4 bytes): A signed integer that specifies the value for this property.
        if fcomplex:
            self.op = len(op)
        else:
            self.op = op

    def get_len(self):
        return 0x6 + (self.opid.is_complex() and self.op or 0)

    def get(self):
        rslt = ''
        rslt += self.opid.get()
        rslt += pack('<L', self.op)
        return rslt

class OfficeArtFOPTEOPID:
    opid = 0x0
    fbid = 0x0
    fcomplex = 0x0

    def __init__(self, opid=0, fbid=0, fcomplex=0):
        self.opid = opid
        self.fbid = fbid
        self.fcomplex = fcomplex

    def is_complex(self):
        return self.fcomplex == 0x1

    def get(self):
        rslt = ''
        rslt += pack('<H', self.opid | self.fbid << 14 | self.fcomplex << 15)
        return rslt

class OfficeArtSplitMenuColorContainer(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x4 # fixed
    _type = 0xF11E
    _len = -1

    def __init__(self):
        self.smca = []
        # dummy
        self.smca.append(0x0800000D)
        self.smca.append(0x0800000C)
        self.smca.append(0x08000017)
        self.smca.append(0x100000F7)

    def _get_len(self):
        return len(self.smca) * 4

    def get(self):
        rslt = self.get_header()
        for x in self.smca:
            rslt += pack('<L', x)
        return rslt




"""
The MSOSPT enumeration, as shown in the following table, specifies the preset shapes and preset
text shape geometries that will be used for a shape. An enumeration of this type is used so that a
custom geometry does not need to be specified but can instead be automatically constructed by the
generating application.
"""
class MSOSPT:
    msosptNotPrimitive = 0x00000000 # A shape that has no preset geometry And is
                                # instead drawn with custom geometry. For
                                # example, freeform shapes that Are drawn by
                                # users fall into this category.
    msosptRectangle = 0x00000001 # A rectangle shape:
    msosptRoundRectangle = 0x00000002 # A rectangle shape with rounded corners:
    msosptEllipse = 0x00000003 # An ellipse shape:
    msosptDiamond = 0x00000004 # A diamond shape:
    msosptIsocelesTriangle = 0x00000005 # An isosceles triangle shape:
    msosptRightTriangle = 0x00000006 # A right triangle shape:
    msosptParallelogram = 0x00000007 # A parallelogram shape:
    msosptTrapezoid = 0x00000008 # A trapezoid shape:
    msosptHexagon = 0x00000009 # A hexagon shape:
    msosptOctagon = 0x0000000A # An octagon shape:
    msosptPlus = 0x0000000B # A plus shape:
    msosptStar = 0x0000000C # A star shape:
    msosptArrow = 0x0000000D # An # Arrow shape:
    msosptThickArrow = 0x0000000E # A value that SHOULD NOT be used.
    msosptHomePlate = 0x0000000F # An irregular pentagon shape:
    msosptCube = 0x00000010 # A cube shape:
    msosptBalloon = 0x00000011 # A speech balloon shape:
    msosptSeal = 0x00000012 # A seal shape:
    msosptArc = 0x00000013 # A curved # Arc shape:
    msosptLine = 0x00000014 # A line shape:
    msosptPlaque = 0x00000015 # A plaque shape:
    msosptCan = 0x00000016 # A cylinder shape:
    msosptDonut = 0x00000017 # A donut shape:
    msosptTextSimple = 0x00000018 # A simple text shape. The text SHOULD<119>
                               # be drawn on # A straight line:
    msosptTextOctagon = 0x00000019 # An octagonal text shape. The text
                               # SHOULD<120> be drawn within # An octagonal
                               # boundary:
    msosptTextHexagon = 0x0000001A # A hexagonal text shape. The text
                               # SHOULD<121> be drawn within # A hexagonal
                               # boundary:
    msosptTextCurve = 0x0000001B # A curved text shape. The text SHOULD<122>
                              # be drawn on # A curved line:
    msosptTextWave = 0x0000001C # A wavy text shape. The text SHOULD<123> be
                             # drawn on # A wavy line:
    msosptTextRing = 0x0000001D # A ringed text shape. The text SHOULD<124>
                             # be drawn within # A semicircular # Arc:
    msosptTextOnCurve = 0x0000001E # A text shape that draws text on # A curve. The
                                # text SHOULD<125> be drawn on # A curved line:
    msosptTextOnRing = 0x0000001F # A text shape that draws text on a ring. The
                               # text SHOULD<126> be drawn on # A
                               # semicircular # Arc:
    msosptStraightConnector1 = 0x00000020 # A straight connector shape:
    msosptBentConnector2 = 0x00000021 # A bent connector shape:
    msosptBentConnector3 = 0x00000022 # A bent connector shape:
    msosptBentConnector4 = 0x00000023 # A bent connector shape:
    msosptBentConnector5 = 0x00000024 # A bent connector shape:
    msosptCurvedConnector2 = 0x00000025 # A curved connector shape:
    msosptCurvedConnector3 = 0x00000026 # A curved connector shape:
    msosptCurvedConnector4 = 0x00000027 # A curved connector shape:
    msosptCurvedConnector5 = 0x00000028 # A curved connector shape:
    msosptCallout1 = 0x00000029 # A callout shape:
    msosptCallout2 = 0x0000002A # A callout shape:
    msosptCallout3 = 0x0000002B # A callout shape:
    msosptAccentCallout1 = 0x0000002C # A callout shape with # A side # Accent:
    msosptAccentCallout2 = 0x0000002D # A callout shape with # A side # Accent:
    msosptAccentCallout3 = 0x0000002E # A callout shape with # A side # Accent:
    msosptBorderCallout1 = 0x0000002F # A callout shape with # A border:
    msosptBorderCallout2 = 0x00000030 # A callout shape with # A border:
    msosptBorderCallout3 = 0x00000031 # A callout shape with # A border:
    msosptAccentBorderCallout1 = 0x00000032 # A callout shape with # A border # And # A side
                                         # Accent:
    msosptAccentBorderCallout2 = 0x00000033 # A callout shape with # A border # And # A side
                                         # Accent:
    msosptAccentBorderCallout3 = 0x00000034 # A callout shape with # A border # And # A side
                                         # Accent:
    msosptRibbon = 0x00000035 # A ribbon shape:
    msosptRibbon2 = 0x00000036 # A ribbon shape:
    msosptChevron = 0x00000037 # A chevron shape:
    msosptPentagon = 0x00000038 # A regular pentagon shape:
    msosptNoSmoking = 0x00000039 # A circle-with-a-slash shape:
    msosptSeal8 = 0x0000003A # A seal shape with eight points:
    msosptSeal16 = 0x0000003B # A seal shape with sixteen points:
    msosptSeal32 = 0x0000003C # A seal shape with thirty-two points:
    msosptWedgeRectCallout = 0x0000003D # A rectangular callout shape:
    msosptWedgeRRectCallout = 0x0000003E # A rectangular callout shape with rounded corners:
    msosptWedgeEllipseCallout = 0x0000003F # An elliptical callout shape:
    msosptWave = 0x00000040 # A wave shape:
    msosptFoldedCorner = 0x00000041 # A rectangular shape with # A folded corner:
    msosptLeftArrow = 0x00000042 # An # Arrow shape that points to the left:
    msosptDownArrow = 0x00000043 # An # Arrow shape that points down:
    msosptUpArrow = 0x00000044 # An # Arrow shape that points up:
    msosptLeftRightArrow = 0x00000045 # An # Arrow shape that points both left # And right:
    msosptUpDownArrow = 0x00000046 # An # Arrow shape that points both down # And up:
    msosptIrregularSeal1 = 0x00000047 # An irregular seal shape:
    msosptIrregularSeal2 = 0x00000048 # An irregular seal shape:
    msosptLightningBolt = 0x00000049 # A lightning bolt shape:
    msosptHeart = 0x0000004A # A heart shape:
    msosptPictureFrame = 0x0000004B # A frame shape:
    msosptQuadArrow = 0x0000004C # A shape that has arrows pointing down, left, right, And up:
    msosptLeftArrowCallout = 0x0000004D # A callout shape that has an arrow pointing to the left:
    msosptRightArrowCallout = 0x0000004E # A callout shape that has # An # Arrow pointing to the right:
    msosptUpArrowCallout = 0x0000004F # A callout shape that has # An # Arrow pointing up:
    msosptDownArrowCallout = 0x00000050 # A callout shape that has # An # Arrow pointing down:
    msosptLeftRightArrowCallout = 0x00000051 # A callout shape that has # Arrows pointing both
                                         # left # And right:
    msosptUpDownArrowCallout = 0x00000052 # A callout shape that has # Arrows pointing both
                                      # down # And up:
    msosptQuadArrowCallout = 0x00000053 # A callout shape that has # Arrows pointing down,
                                    # left, right, # And up:
    msosptBevel = 0x00000054 # A beveled rectangle shape:
    msosptLeftBracket = 0x00000055 # An opening bracket shape:
    msosptRightBracket = 0x00000056 # A closing bracket shape:
    msosptLeftBrace = 0x00000057 # An opening brace shape:
    msosptRightBrace = 0x00000058 # A closing brace shape:
    msosptLeftUpArrow = 0x00000059 # An # Arrow shape that points both left # And up:
    msosptBentUpArrow = 0x0000005A # A bent # Arrow shape that has its base on the left
                                   # and that points up:
    msosptBentArrow = 0x0000005B # A curved # Arrow shape that has its base on the
                              # bottom # And that points to the right:
    msosptSeal24 = 0x0000005C # A seal shape with twenty-four points:
    msosptStripedRightArrow = 0x0000005D # A striped # Arrow shape that points to the right:
    msosptNotchedRightArrow = 0x0000005E # A notched # Arrow shape that points to the right:
    msosptBlockArc = 0x0000005F # A semicircular # Arc shape:
    msosptSmileyFace = 0x00000060 # A smiling face shape:
    msosptVerticalScroll = 0x00000061 # A scroll shape that is vertically opened:
    msosptHorizontalScroll = 0x00000062 # A scroll shape that is horizontally opened:
    msosptCircularArrow = 0x00000063 # A semicircular # Arrow shape:
    msosptNotchedCircularArrow = 0x00000064 # A value that SHOULD NOT be used.
    msosptUturnArrow = 0x00000065 # A semicircular # Arrow shape that has # A straight tail:
    msosptCurvedRightArrow = 0x00000066 # An # Arrow shape that curves to the right:
    msosptCurvedLeftArrow = 0x00000067 # An # Arrow shape that curves to the left:
    msosptCurvedUpArrow = 0x00000068 # An # Arrow shape that curves upward:
    msosptCurvedDownArrow = 0x00000069 # An # Arrow shape that curves downward:
    msosptCloudCallout = 0x0000006A # A cloud-shaped callout:
    msosptEllipseRibbon = 0x0000006B # An elliptical ribbon shape:
    msosptEllipseRibbon2 = 0x0000006C # An elliptical ribbon shape:
    msosptFlowChartProcess = 0x0000006D # A process shape for flowcharts:
    msosptFlowChartDecision = 0x0000006E # A decision shape for flowcharts:
    msosptFlowChartInputOutput = 0x0000006F # An input-output shape for flowcharts:
    msosptFlowChartPredefinedProcess = 0x00000070 # A predefined process shape for flowcharts:
    msosptFlowChartInternalStorage = 0x00000071 # An internal storage shape for flowcharts:
    msosptFlowChartDocument = 0x00000072 # A document shape for flowcharts:
    msosptFlowChartMultidocument = 0x00000073 # A multiple-document shape for flowcharts:
    msosptFlowChartTerminator = 0x00000074 # A terminator shape for flowcharts:
    msosptFlowChartPreparation = 0x00000075 # A preparation shape for flowcharts:
    msosptFlowChartManualInput = 0x00000076 # A manual input shape for flowcharts:
    msosptFlowChartManualOperation = 0x00000077 # A manual operation shape for flowcharts:
    msosptFlowChartConnector = 0x00000078 # A connector shape for flowcharts:
    msosptFlowChartPunchedCard = 0x00000079 # A punched card shape for flowcharts:
    msosptFlowChartPunchedTape = 0x0000007A # A punched tape shape for flowcharts:
    msosptFlowChartSummingJunction = 0x0000007B # A summing junction shape for flowcharts:
    msosptFlowChartOr = 0x0000007C # An OR shape for flowcharts:
    msosptFlowChartCollate = 0x0000007D # A collation shape for flowcharts:
    msosptFlowChartSort = 0x0000007E # A sorting shape for flowcharts:
    msosptFlowChartExtract = 0x0000007F # An extraction shape for flowcharts:
    msosptFlowChartMerge = 0x00000080 # A merging shape for flowcharts:
    msosptFlowChartOfflineStorage = 0x00000081 # An offline storage shape for flowcharts:
    msosptFlowChartOnlineStorage = 0x00000082 # An online storage shape for flowcharts:
    msosptFlowChartMagneticTape = 0x00000083 # A magnetic tape shape for flowcharts:
    msosptFlowChartMagneticDisk = 0x00000084 # A magnetic disk shape for flowcharts:
    msosptFlowChartMagneticDrum = 0x00000085 # A magnetic drum shape for flowcharts:
    msosptFlowChartDisplay = 0x00000086 # A display shape for flowcharts:
    msosptFlowChartDelay = 0x00000087 # A delay shape for flowcharts:
    msosptTextPlainText = 0x00000088 # A plain text shape:
    msosptTextStop = 0x00000089 # An octagonal text shape:
    msosptTextTriangle = 0x0000008A # A triangular text shape that points upward:
    msosptTextTriangleInverted = 0x0000008B # A triangular text shape that points downward:
    msosptTextChevron = 0x0000008C # A chevron text shape that points upward:
    msosptTextChevronInverted = 0x0000008D # A chevron text shape that points downward:
    msosptTextRingInside = 0x0000008E # A circular text shape, in which reading the text
                                 #  is like reading # An inscription on the inside of # A
                                 # ring:
    msosptTextRingOutside = 0x0000008F # A circular text shape, in which reading the text
                                 #   is like reading # An inscription on the outside of # A
                                 #  ring:
    msosptTextArchUpCurve = 0x00000090 # An upward-arching curved text shape:
    msosptTextArchDownCurve = 0x00000091 # A downward-arching curved text shape:
    msosptTextCircleCurve = 0x00000092 # A circular text shape:
    msosptTextButtonCurve = 0x00000093 # A text shape that resembles # A button:
    msosptTextArchUpPour = 0x00000094 # An upward-arching text shape:
    msosptTextArchDownPour = 0x00000095 # A downward-arching text shape:
    msosptTextCirclePour = 0x00000096 # A circular text shape:
    msosptTextButtonPour = 0x00000097 # A text shape that resembles # A button:
    msosptTextCurveUp = 0x00000098 # An upward-curving text shape:
    msosptTextCurveDown = 0x00000099 # A downward-curving text shape:
    msosptTextCascadeUp = 0x0000009A # A cascading text shape that points up:
    msosptTextCascadeDown = 0x0000009B # A cascading text shape that points down:
    msosptTextWave1 = 0x0000009C # A wavy text shape:
    msosptTextWave2 = 0x0000009D # A wavy text shape:
    msosptTextWave3 = 0x0000009E # A wavy text shape:
    msosptTextWave4 = 0x0000009F # A wavy text shape:
    msosptTextInflate = 0x000000A0 # A text shape that vertically expands in the
                                # middle:
    msosptTextDeflate = 0x000000A1 # A text shape that vertically shrinks in the
                                # middle:
    msosptTextInflateBottom = 0x000000A2 # A text shape that expands downward in the
                                      # middle:
    msosptTextDeflateBottom = 0x000000A3 # A text shape that shrinks upward in the
                                      # middle:
    msosptTextInflateTop = 0x000000A4 # A text shape that expands upward in the
                                   # middle:
    msosptTextDeflateTop = 0x000000A5 # A text shape that shrinks downward in the
                                   # middle:
    msosptTextDeflateInflate = 0x000000A6 # A text shape in which the lower lines expand
                                      # upward, And the upper lines shrink to
                                      # compensate:
    msosptTextDeflateInflateDeflate = 0x000000A7 # A text shape in which the lines in the center
                                            #  vertically expand, # And the upper # And lower
                                            # lines shrink to compensate:
    msosptTextFadeRight = 0x000000A8 # A text shape that vertically shrinks on the right
                                #  side:
    msosptTextFadeLeft = 0x000000A9 # A text shape that vertically shrinks on the left
                                # side:
    msosptTextFadeUp = 0x000000AA # A text shape that horizontally shrinks on the
                               # top:
    msosptTextFadeDown = 0x000000AB # A text shape that horizontally shrinks on the
                               # bottom:
    msosptTextSlantUp = 0x000000AC # An upward-slanted text shape:
    msosptTextSlantDown = 0x000000AD # A downward-slanted text shape:
    msosptTextCanUp = 0x000000AE # A text shape that is curved upward # As if being
                              # read on the side of # A can:
    msosptTextCanDown = 0x000000AF # A text shape that is curved downward # As if
                                    # being read on the side of # A can:
    msosptFlowChartAlternateProcess = 0x000000B0 # An # Alternate process shape for flowcharts:
    msosptFlowChartOffpageConnector = 0x000000B1 # An off-page connector shape for flowcharts:
    msosptCallout90 = 0x000000B2 # A callout shape:
    msosptAccentCallout90 = 0x000000B3 # A callout shape:
    msosptBorderCallout90 = 0x000000B4 # A callout shape with # A border:
    msosptAccentBorderCallout90 = 0x000000B5 # A callout shape with # A border:
    msosptLeftRightUpArrow = 0x000000B6 # A shape that has # Arrows pointing left, right,
                                     # And up:
    msosptSun = 0x000000B7 # A sun shape:
    msosptMoon = 0x000000B8 # A moon shape:
    msosptBracketPair = 0x000000B9 # A shape that is enclosed in brackets:
    msosptBracePair = 0x000000BA # A shape that is enclosed in braces:
    msosptSeal4 = 0x000000BB # A seal shape with four points:
    msosptDoubleWave = 0x000000BC # A double wave shape:
    msosptActionButtonBlank = 0x000000BD # A blank button shape:
    msosptActionButtonHome = 0x000000BE # A home button shape:
    msosptActionButtonHelp = 0x000000BF # A help button shape:
    msosptActionButtonInformation = 0x000000C0 # An information button shape:
    msosptActionButtonForwardNext = 0x000000C1 # A forward or next button shape:
    msosptActionButtonBackPrevious = 0x000000C2 # A back or previous button shape:
    msosptActionButtonEnd = 0x000000C3 # An end button shape:
    msosptActionButtonBeginning = 0x000000C4 # A beginning button shape:
    msosptActionButtonReturn = 0x000000C5 # A return button shape:
    msosptActionButtonDocument = 0x000000C6 # A document button shape:
    msosptActionButtonSound = 0x000000C7 # A sound button shape:
    msosptActionButtonMovie = 0x000000C8 # A movie button shape:
    msosptHostControl = 0x000000C9 # A value that SHOULD NOT be used.
    msosptTextBox = 0x000000CA # A text box shape:








class OfficeArtDgContainer(ODrawRecordBase):
    _ver = 0xF
    _ins = 0x0
    _type = 0xF002
    _len = -1

    sheet_id = 0

    def __init__(self, sheet_id):
        self.sheet_id = sheet_id
        # drawingData (16 bytes): An OfficeArtFDG record, as defined in section 2.2.49, that specifies
        # the shape count, drawing identifier, and shape identifier of the last shape in this drawing.
        self.drawing_data = OfficeArtFDG()
        self.drawing_data._ins = self.sheet_id + 1
        self.drawing_data.spid_cur &= 0x0001
        self.drawing_data.spid_cur |= (self.sheet_id + 1) << 0x08
        # regroupItems (variable): An OfficeArtFRITContainer record, as defined in section 2.2.41,
        # that specifies a container for the table of group (4) identifiers for regrouping ungrouped
        # shapes.
        self.regroup_items = OfficeArtFRITContainer()
        # groupShape (variable): An OfficeArtSpgrContainer record, as defined in section 2.2.16, that
        # specifies a container for groups (4) of shapes.
        self.group_shape = OfficeArtSpgrContainer(sheet_id)
        # shape (variable): An OfficeArtSpContainer record, as defined in section 2.2.14, that specifies
        # a container for the shapes that are not contained in a group (4).
        self.shape = UndefinedBase() #OfficeArtSpContainer(sheet_id)
        # deletedShapes (variable): An array of OfficeArtSpgrContainerFileBlock records, as defined
        # in section 2.2.17, that specifies the deleted shapes. For more information, see section 2.2.37.
        # The array continues if the rh.recType field of the OfficeArtSpgrContainerFileBlock record,
        # as defined in section 2.2.17, equals 0xF003 or 0xF004. This array MAY<2> exist.
        self.deleted_shapes = []
        # solvers2 (variable): An OfficeArtSolverContainer record, as defined in section 2.2.18, that
        # specifies a container for the rules (1) that are applicable to the shapes contained in this
        # drawing.
        self.solvers2 = OfficeArtSolverContainer()

    def _get_len(self):
        rslt = 0
        rslt += self.drawing_data.get_full_len()
        rslt += self.regroup_items.get_full_len()
        rslt += self.group_shape.get_full_len()
        rslt += self.shape.get_full_len()
        rslt += sum([x.get_full_len() for x in self.deleted_shapes])
        rslt += self.solvers2.get_full_len()
        #print self.child_count, rslt, '+', self.child_count * 0xA0
        #if self.child_count > 0:
        #    rslt += (self.child_count - 1) * 0xA0
        #print '=', rslt
        return rslt

    def add(self):
        self.drawing_data.insert()
        self.group_shape.add()

    def get_count(self):
        return self.drawing_data.csp

    def insert(self, pic_id, shape_id, pic_type, *coordinates):
        #self.drawing_data.insert()
        self.group_shape.insert(pic_id, shape_id, pic_type, *coordinates)

    def get(self):
        if self.get_count() == 0:
            return ''
        rslt = self.get_header()
        rslt += self.drawing_data.get()
        rslt += self.regroup_items.get()
        rslt += self.group_shape.get()
        rslt += self.shape.get()
        rslt += ''.join([x.get() for x in self.deleted_shapes])
        rslt += self.solvers2.get()
        return rslt

class OfficeArtFDG(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x0 # the drawing identifier 
    _type = 0xF008
    _len = 0x08

    def __init__(self):
        # csp (4 bytes): An unsigned integer that specifies the number of shapes in this drawing.
        # indexed / stored by sheet
        self.csp = 0x01
        # spidCur (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the shape
        # identifier of the last shape in this drawing.
        self.spid_cur = 0x0400

    def insert(self):
        self.csp += 1
        self.spid_cur += 1

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.csp)
        rslt += pack('<L', self.spid_cur)
        return rslt

class OfficeArtSpgrContainerFileBlock(ODrawRecordBase):
    pass

class OfficeArtSpgrContainer(OfficeArtSpgrContainerFileBlock):
    _ver = 0xF
    _ins = 0x0 # the drawing identifier 
    _type = 0xF003
    _len = -1

    sheet_id = 0

    def __init__(self, sheet_id):
        # rgfb (variable): An array of OfficeArtSpgrContainerFileBlock records, as defined in section
        # 2.2.17, that specifies the groups (4) or shapes that are contained within this group (4).
        self.rgfbs = []  # per file

        self.sheet_id = sheet_id
        self.child_count = 0
        
        # make default
        sp = OfficeArtSpContainer(self.sheet_id, default=True)
        self.rgfbs.append(sp)

    def insert(self, pic_id, shape_id, pic_type, *coordinates):
        #cnt = len(self.rgfbs)
        sp = OfficeArtSpContainer(self.sheet_id)
        sp.insert(pic_id, shape_id, pic_type, *coordinates)
        self.rgfbs.append(sp)

    def add(self):
        self.child_count += 1

    def _get_len(self):
        rslt = sum([x.get_full_len() for x in self.rgfbs])
        if self.child_count > 0:
            rslt += (self.child_count - 1) * 0xA0
        return rslt

    def get(self):
        rslt = self.get_header()
        rslt += ''.join([x.get() for x in self.rgfbs])
        return rslt

class OfficeArtSpContainer(OfficeArtSpgrContainerFileBlock):
    _ver = 0xF
    _ins = 0x0 # fixed 
    _type = 0xF004
    _len = -1

    sheet_id = 0

    def __init__(self, sheet_id, default=False):
        self.sheet_id = sheet_id
        # shapeGroup (24 bytes): An OfficeArtFSPGR record, as defined in section 2.2.38, that
        # specifies the coordinate system of the group shape. The anchors of the child shape are
        # expressed in this coordinate system. This record’s container MUST be a group shape.
        self.shape_group = default and OfficeArtFSPGR() or UndefinedBase()
        # shapeProp (16 bytes): An OfficeArtFSP record, as defined in section 2.2.40, that specifies an
        # instance of a shape.
        self.shape_prop = OfficeArtFSP(shape_type=0, spid=(self.sheet_id + 1) << 10, options='ac')
        # deletedShape (12 bytes): An OfficeArtFPSPL record, as defined in section 2.2.37, that
        # specifies the former hierarchical position of the containing object. This record’s container
        # MUST be a deleted shape. For more information, see OfficeArtFPSPL.
        self.deleted_shape = UndefinedBase() #OfficeArtFPSPL()
        # shapePrimaryOptions (variable): An OfficeArtFOPT record, as defined in section 2.2.9, that
        # specifies the properties of this shape that do not contain default values.
        self.pri_opt = OfficeArtFOPT()
        # shapeSecondaryOptions1 (variable): An OfficeArtSecondaryFOPT record, as defined in
        # section 2.2.10, that specifies the properties of this shape that do not contain default values.
        self.sec_opt1 = OfficeArtSecondaryFOPT()
        # shapeTertiaryOptions1 (variable): An OfficeArtTertiaryFOPT record, as defined in section
        # 2.2.11, that specifies the properties of this shape that do not contain default values.
        self.ter_opt1 = OfficeArtTertiaryFOPT()
        # childAnchor (24 bytes): An OfficeArtChildAnchor record, as defined in section 2.2.39, that
        # specifies the anchor for this shape. This record’s container MUST be a member of a group (4)
        # of shapes.
        self.child_anchor = UndefinedBase() #OfficeArtChildAnchor()
        # clientAnchor (variable): An OfficeArtClientAnchor record as specified by the host
        # application.
        self.client_anchor = default and UndefinedBase() or OfficeArtClientAnchor()
        # clientData (variable): An OfficeArtClientData record as specified by the host application.
        self.client_data = default and UndefinedBase() or OfficeArtClientData()
        # clientTextbox (variable): An OfficeArtClientTextbox record as specified by the host
        # application.
        self.client_textbox = default and UndefinedBase() or OfficeArtClientTextbox()
        # shapeSecondaryOptions2 (variable): An OfficeArtSecondaryFOPT record that specifies the
        # properties of this shape that do not contain default values. This field MUST NOT exist if
        # shapeSecondaryOptions1 exists.
        self.sec_opt2 = default and UndefinedBase() or OfficeArtSecondaryFOPT()
        # shapeTertiaryOptions2 (variable): An OfficeArtTertiaryFOPT record, as defined in section
        # 2.2.11, that specifies the properties of this shape that do not contain default values. This field
        # MUST NOT exist if shapeTertiaryOptions1 exists.
        self.ter_opt2 = default and UndefinedBase() or OfficeArtTertiaryFOPT()

    def insert(self, pic_id, shape_id, pic_type, *coordinates):
        self.shape_prop._ins = MSOSPT.msosptPictureFrame # 4B
        self.shape_prop.spid = (self.sheet_id + 1) << 10 | shape_id #pic_id
        self.shape_prop.options = 'jl'

        self.pri_opt.insert(0x007F, 0x00800080, 0, 0)
        self.pri_opt.insert(0x0085, 0x00000002, 0, 0)
        self.pri_opt.insert(0x0087, 0x00000001, 0, 0)
        self.pri_opt.insert(0x0104, pic_id, 1, 0)
        self.pri_opt.insert(0x0180, 0x00000003, 0, 0)
        self.pri_opt.insert(0x01BF, 0x00100000, 0, 0)
        self.pri_opt.insert(0x01C0, 0x00000000, 0, 0)
        self.pri_opt.insert(0x01C2, 0x00FFFFFF, 0, 0)
        #self.pri_opt.insert(0x01CB, 0x00002535, 0, 0)
        self.pri_opt.insert(0x01D6, 0x00000002, 0, 0)
        self.pri_opt.insert(0x01FF, 0x00090000, 0, 0)
        self.pri_opt.insert(0x023F, 0x00020000, 0, 0)
        # test zoom 50%
        #self.pri_opt.insert(0x07C0, 0x000001f4, 0, 0)
        #self.pri_opt.insert(0x07C4, 0x00000001, 0, 0)

        # FIXME: if the image name is not 16bits long, excel will complain when opening the book.
        #        this is weird but before it's fixed we just ensure the image name to 16bits.
        prefix = 'Graphics'
        img_name = str(pic_id)
        pre_len = 9 - len(str(pic_id))
        img_name = '%s %s\0' % (prefix[:pre_len], str(pic_id))
        img_name = ''.join([x + '\0' for x in img_name])
        self.pri_opt.insert(0x0380, img_name, 1, 1)

        self.client_anchor.set_coordinates(*coordinates)

    def _get_len(self):
        rslt = 0
        rslt += self.shape_group.get_full_len()
        rslt += self.shape_prop.get_full_len()
        rslt += self.deleted_shape.get_full_len()
        rslt += self.pri_opt.get_full_len()
        rslt += self.sec_opt1.get_full_len()
        rslt += self.ter_opt1.get_full_len()
        rslt += self.child_anchor.get_full_len()
        rslt += self.client_anchor.get_full_len()
        rslt += self.client_data.get_full_len()
        rslt += self.client_textbox.get_full_len()
        rslt += self.sec_opt2.get_full_len()
        rslt += self.ter_opt2.get_full_len()

        """
        print 'shape group', self.shape_group.get_full_len()
        print 'shape prop', self.shape_prop.get_full_len()
        print 'deleted shape', self.deleted_shape.get_full_len()
        print 'pri opt', self.pri_opt.get_full_len()
        print 'sec opt1', self.sec_opt1.get_full_len()
        print 'ter opt1', self.ter_opt1.get_full_len()
        print 'child anchor', self.child_anchor.get_full_len()
        print 'client anchor',  self.client_anchor.get_full_len()
        print 'client data', self.client_data.get_full_len()
        print 'client textbox',  self.client_textbox.get_full_len()
        print 'sec opt2', self.sec_opt2.get_full_len()
        print 'ter opt2', self.ter_opt2.get_full_len()
        print 'total', rslt
        """
        return rslt

    def get(self):
        rslt = self.get_header()
        rslt += self.shape_group.get()
        rslt += self.shape_prop.get()
        rslt += self.deleted_shape.get()
        rslt += self.pri_opt.get()
        rslt += self.sec_opt1.get()
        rslt += self.ter_opt1.get()
        rslt += self.child_anchor.get()
        rslt += self.client_anchor.get()
        rslt += self.client_data.get()
        rslt += self.client_textbox.get()
        rslt += self.sec_opt2.get()
        rslt += self.ter_opt2.get()
        return rslt

class OfficeArtSolverContainer(ODrawRecordBase):
    _ver = 0xF
    _ins = 0x0 # count
    _type = 0xF005
    _len = -1

    def __init__(self):
        # rgfb (variable): An array of OfficeArtSolverContainerFileBlock records, as defined in section
        # 2.2.19, specifying a collection of rules (1) that are applicable to the shapes contained in an
        # OfficeArtDgContainer record, as defined in section 2.2.13.
        self.rgfbs = []

    def _get_len(self):
        return sum([x.get_full_len() for x in self.rgfbs])

    def get(self):
        if not self.rgfbs:
            return ''
        rslt = self.get_header()
        rslt += ''.join([x.get() for x in self.rgfbs])
        return rslt

class OfficeArtSolverContainerFileBlock(ODrawRecordBase):
    pass

class OfficeArtFConnectorRule(OfficeArtSolverContainerFileBlock):
    _ver = 0x1
    _ins = 0x0 # fixed
    _type = 0xF012
    _len = 0x18

    def __init__(self):
        # ruid (4 bytes): An unsigned integer that specifies the identifier of this rule (1).
        self.ruid = 0x0
        # spidA (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the
        # identifier of the shape where the connector shape starts.
        self.spida = 0x0
        # spidB (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the identifier
        # of the shape where the connector shape ends.
        self.spidb = 0x0
        # spidC (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the identifier
        # of the connector shape.
        self.spidc = 0x0
        # cptiA (4 bytes): An unsigned integer that specifies the connection site index of the shape
        # where the connector shape starts. If the shape is available, this value MUST be within its
        # range of valid connection site indexes. Otherwise, this value is ignored.
        self.cptia = 0x0
        # cptiB (4 bytes): An unsigned integer that specifies the connection site index of the shape where
        # the connector shape ends. If the shape is available, this value MUST be within its range of
        # valid connection site indexes. Otherwise, this value is ignored.
        self.cptib = 0x0

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.ruid)
        rslt += pack('<L', self.spida)
        rslt += pack('<L', self.spidb)
        rslt += pack('<L', self.spidc)
        rslt += pack('<L', self.cptia)
        rslt += pack('<L', self.cptib)
        return rslt

class OfficeArtFArcRule(OfficeArtSolverContainerFileBlock):
    _ver = 0x0
    _ins = 0x0 # fixed
    _type = 0xF014
    _len = 0x08

    def __init__(self):
        # ruid (4 bytes): An unsigned integer that specifies the identifier of this arc rule (1).
        self.ruid = 0x0
        # spid (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the identifier
        # of the arc shape.
        self.spid = 0x0

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.ruid)
        rslt += pack('<L', self.spid)
        return rslt
 
class OfficeArtFCalloutRule(OfficeArtSolverContainerFileBlock):
    _ver = 0x0
    _ins = 0x0 # fixed
    _type = 0xF017
    _len = 0x08

    def __init__(self):
        # ruid (4 bytes): An unsigned integer that specifies the identifier of this arc rule (1).
        self.ruid = 0x0
        # spid (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the identifier
        # of the arc shape.
        self.spid = 0x0

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.ruid)
        rslt += pack('<L', self.spid)
        return rslt
 
class OfficeArtFRITContainer(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x0 # count
    _type = 0xF118
    _len = -1

    def __init__(self):
        # rgfrit (variable): An array of OfficeArtFRIT records, as defined in section 2.2.42, that specifies
        # the table of group (4) identifiers. The size of the array MUST equal the value of
        # rh.recInstance.
        self.rgfrits = []

    def _get_len(self):
        return sum([x.get_full_len() for x in self.rgfrits])

    def get(self):
        if not self.rgfrits:
            return ''
        rslt = self.get_header()
        rslt += ''.join([x.get() for x in self.rgfrits])
        return rslt

class OfficeArtFRIT:
    def __init__(self, new, old):
        # fridNew (2 bytes): A FRID structure, as defined in section 2.1.3, specifying the last group (4)
        # identifier of the shape before ungrouping. The value of fridNew MUST be greater than the
        # value of fridOld.
        self.frid_new = new
        # fridOld (2 bytes): A FRID structure, as defined in section 2.1.3, specifying the second-to-last
        # group (4) identifier of the shape before ungrouping. This value MUST be 0x0000 if a second-
        # to-last group (4) does not exist.
        self.frid_old = old

    def get_len(self):
        return 0x04

    def get(self):
        rslt = ''
        rslt += pack('<H', self.frid_new)
        rslt += pack('<H', self.frid_old)
        return rslt

class OfficeArtFSPGR(ODrawRecordBase):
    _ver = 0x1
    _ins = 0x0 # fixed
    _type = 0xF009
    _len = 0x10

    def __init__(self, l=0, t=0, r=0, b=0):
        # xLeft (4 bytes): A signed integer that specifies the left boundary of the coordinate system of
        # the group (4).
        self.xleft = l
        # yTop (4 bytes): A signed integer that specifies the top boundary of the coordinate system of the
        # group (4).
        self.ytop = t
        # xRight (4 bytes): A signed integer that specifies the right boundary of the coordinate system of
        # the group (4).
        self.xright = r
        # yBottom (4 bytes): A signed integer that specifies the bottom boundary of the coordinate
        # system of the group (4).
        self.ybottom = b

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.xleft)
        rslt += pack('<L', self.ytop)
        rslt += pack('<L', self.xright)
        rslt += pack('<L', self.ybottom)
        return rslt

class OfficeArtChildAnchor(OfficeArtFSPGR):
    _ver = 0x0
    _ins = 0x0 # fixed
    _type = 0xF00F
    _len = 0x10

class UndefinedBase(ODrawRecordBase):
    def get_len(self):
        return 0

    def get(self):
        return ''

class OfficeArtClientAnchor(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x0 # fixed
    _type = 0xF010
    _len = 0x12

    def __init__(self):
        # A - fMove (1 bit): A bit that specifies whether the shape will be kept intact when the cells are moved.
        self.fmove = 0x0
        # B - fSize (1 bit): A bit that specifies whether the shape will be kept intact when the cells are resized. If fMove is 1, the value MUST be 1.
        self.fsize = 0x0
        # C - reserved1 (1 bit): MUST be 0 and MUST be ignored.
        # D - reserved2 (1 bit): MUST be 0 and MUST be ignored
        # E - reserved3 (1 bit): MUST be 0 and MUST be ignored.
        # unused (11 bits): Undefined and MUST be ignored.
        self.unused = 0x0
        # colL (2 bytes): A Col256U that specifies the column of the cell under the top left corner of the bounding rectangle of the shape.
        self.cl = 0x0
        # dxL (2 bytes): A signed integer that specifies the x coordinate of the top left corner of the bounding rectangle relative to the corner of the underlying cell. The value is expressed as 1024th’s of that cell’s width.
        self.xl = 0x0
        # rwT (2 bytes): A RwU that specifies the row of the cell under the top left corner of the bounding rectangle of the shape.
        self.rt = 0x0
        # dyT (2 bytes): A signed integer that specifies the y coordinate of the top left corner of the bounding rectangle relative to the corner of the underlying cell. The value is expressed as 1024th’s of that cell’s height. 
        self.yt = 0x0
        # colR (2 bytes): A Col256U that specifies the column of the cell under the bottom right corner of the bounding rectangle of the shape.
        self.cr = 0x0
        # dxR (2 bytes): A signed integer that specifies the x coordinate of the bottom right corner of the bounding rectangle relative to the corner of the underlying cell. The value is expressed as 1024th’s of that cell’s width. 
        self.xr = 0x0
        # rwB (2 bytes): A RwU that specifies the row of the cell under the bottom right corner of the bounding rectangle of the shape.
        self.rb = 0x0
        # dyB (2 bytes): A signed integer that specifies the y coordinate of the bottom right corner of the bounding rectangle relative to the corner of the underlying cell. The value is expressed as 1024th’s of that cell’s height. 
        self.yb = 0x0
        self.set_default()
        """
        self.data4 = 0x03360007
        self.data5 = 0x003B001C
        """

    def set_default(self):
        self.cr, self.rb = 0x07, 0x05
        self.xr, self.yb = 0x01, 0x01

    def set_coordinates(self, *coordinates):
        #print coordinates
        self.cl, self.xl, self.rt, self.yt, self.cr, self.xr, self.rb, self.yb = coordinates

    def get(self):
        rslt = self.get_header()
        rslt += pack('<H', self.fmove | self.fsize << 1 | self.unused << 2)
        rslt += pack('<4H', self.cl, self.xl, self.rt, self.yt)
        rslt += pack('<4H', self.cr, self.xr, self.rb, self.yb)
        return rslt

class OfficeArtClientData(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x0 # fixed
    _type = 0xF011
    _len = 0x0

    def get_full_len(self):
        return 0x8

    def get(self):
        rslt = self.get_header()
        return rslt

class OfficeArtClientTextbox(UndefinedBase):
    pass

class OfficeArtFSP(ODrawRecordBase):
    _ver = 0x2
    _ins = 0x0 # A signed value that specifies the shape type and that MUST be an MSOSPT enumeration value, as defined in section 2.4.24.
    _type = 0xF00A
    _len = 0x08

    def __init__(self, shape_type, spid, options=''):
        self._ins = shape_type

        # spid (4 bytes): An MSOSPID structure, as defined in section 2.1.2, that specifies the identifier
        #                 of this shape.
        # related to sheet index
        # sheet1: 0x0400(fixed) 0x0401 0x0402 ...
        # sheet2: 0x0800(fixed) 0x0801 0x0802 ...
        # sheet3: 0x0C00(fixed) 0x0C01 0x0C02 ...
        # the fixed part will always written even if no image in that sheet
        self.spid = spid 

        # A - fGroup (1 bit): A bit that specifies whether this shape is a group shape.
        # B - fChild (1 bit): A bit that specifies whether this shape is a child shape.
        # C - fPatriarch (1 bit): A bit that specifies whether this shape is the topmost group shape. Each
        #                         drawing contains one topmost group shape.
        # D - fDeleted (1 bit): A bit that specifies whether this shape has been deleted.
        # E - fOleShape (1 bit): A bit that specifies whether this shape is an OLE object.
        # F - fHaveMaster (1 bit): A bit that specifies whether this shape has a valid master in the
        #                          hspMaster property, as defined in section 2.3.2.1.
        # G - fFlipH (1 bit): A bit that specifies whether this shape is horizontally flipped.
        # H - fFlipV (1 bit): A bit that specifies whether this shape is vertically flipped.
        # I - fConnector (1 bit): A bit that specifies whether this shape is a connector shape.
        # J - fHaveAnchor (1 bit): A bit that specifies whether this shape has an anchor.
        # K - fBackground (1 bit): A bit that specifies whether this shape is a background shape.
        # L - fHaveSpt (1 bit): A bit that specifies whether this shape has a shape type property.
        self.options = options
        # unused1 (20 bits): A value that is undefined and MUST be ignored.
        self.unused1 = 0x0

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.spid)
        option = 0x0
        #print self.options
        if 'a' in self.options: option |= 0x1
        if 'b' in self.options: option |= 0x1 << 1
        if 'c' in self.options: option |= 0x1 << 2
        if 'd' in self.options: option |= 0x1 << 3
        if 'e' in self.options: option |= 0x1 << 4
        if 'f' in self.options: option |= 0x1 << 5
        if 'g' in self.options: option |= 0x1 << 6
        if 'h' in self.options: option |= 0x1 << 7
        if 'i' in self.options: option |= 0x1 << 8
        if 'j' in self.options: option |= 0x1 << 9
        if 'k' in self.options: option |= 0x1 << 10
        if 'l' in self.options: option |= 0x1 << 11
        rslt += pack('<L', option)
        return rslt

class OfficeArtFPSPL(ODrawRecordBase):
    _ver = 0x0
    _ins = 0x0 # fixed
    _type = 0xF11D
    _len = 0x04

    def __init__(self, spid, flast):
        # spid (30 bits): An MSOSPID structure, as defined in section 2.1.2, that specifies another shape
        #                 or group (4) of shapes that is contained in the same OfficeArtDgContainer record, as
        #                 defined in section 2.2.13. This other object contains an OfficeArtFSP record, as defined in
        #                 section 2.2.40, with an equivalently valued spid field.
        self.spid = spid
        # A - reserved1 (1 bit): A value that MUST be zero and MUST be ignored.
        self.reserved1 = 0x0
        # B - fLast (1 bit): A bit that specifies the ordering of this record’s containing object and the
        #                    object that is specified by spid. The following table specifies the meaning of each value for
        #                    this bit.
        self.flast = flast
        # Value 0 This record’s containing object was formerly antecedent to the object that is referenced by
        #          spid, in the container directly containing that object.
        # Value 1 This record’s containing object was formerly subsequent to the object that is referenced by
        #          spid, in the container directly containing that object.

    def get(self):
        rslt = self.get_header()
        rslt += pack('<L', self.spid | self.reserved << 30 | self.flast << 31)
        return rslt
