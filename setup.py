#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os.path, glob

from setuptools import setup

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
}

setup(**params)
