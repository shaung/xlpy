#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os, glob

#from setuptools import setup
from distutils.core import setup
from distutils.extension import Extension

try:
    from Cython.Distutils import build_ext
except ImportError:
    use_cython = False
else:
    use_cython = True

cmdclass = {}
ext_modules = []

if use_cython:
    ext_modules = [
        Extension("xlpy.xlwt.cell", [ "xlpy/xlwt/cell.pyx", "xlpy/xlwt/cell.pxd" ]),
        Extension("xlpy.xlwt.row", [ "xlpy/xlwt/row.pyx" ]),
        Extension("xlpy.xlwt.worksheet", [ "xlpy/xlwt/worksheet.pyx" ]),
    ]
    cmdclass.update({ 'build_ext': build_ext })
else:
    ext_modules = [
        Extension("xlpy.xlwt.cell", [ "xlpy/xlwt/cell.c" ]),
        Extension("xlpy.xlwt.row", [ "xlpy/xlwt/row.c" ]),
        Extension("xlpy.xlwt.worksheet", [ "xlpy/xlwt/worksheet.c" ]),
    ]


__VERSION__ = '0.0.1.0'

params = {
    'name': 'xlpy',
    'version': __VERSION__,
    'description': 'edit excel97 files with python',
    'author': 'Shaung',
    'author_email': 'shaun.geng@gmail.com',
    'url': 'http://github.com/shaung/xlpy/',
    'packages':[
        'xlpy',
        'xlpy.xlrd',
        'xlpy.xlwt',
        'xlpy.xlutils',
    ],
    'license': 'BSD',
    'cmdclass': cmdclass,
    'ext_modules': ext_modules,
}

setup(**params)
