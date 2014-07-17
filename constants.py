# -*- coding: UTF-8 -*-
import os, sys
def resource_path(relative):
    return os.path.join(getattr(sys, '_MEIPASS', os.path.abspath(".")),
                        relative)

#cert_path = resource_path('cacert.pem')

VERSION = 1.1
APPNAME = "pyO2g"
CFGFILE = '%s.cfg' % APPNAME
DBFILE = 'oSync.db'
