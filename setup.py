#!/usr/bin/env python
# coding=cp1251

from distutils.core import setup
import py2exe

setup(
    # The first three parameters are not required, if at least a
    # 'version' is given, then a versioninfo resource is built from
    # them and added to the executables.
    version="0.0.0.1",
    description="MedBlanks",
    name="MedBlanks",
    data_files=[("", ["blank.docx"])],

    # targets to build
    windows=[{"script": "MedBlanks.pyw"}],
    options={"py2exe":
        {
            "includes": ["sip"],
            "dll_excludes": ["MSVCP90.dll", "MSWSOCK.DLL"],
            "dist_dir": "MedBlanks"
        }
    }

)
