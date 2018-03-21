# coding=utf-8
"""
  useage:
    python setup.py py2exe
"""

import glob
from distutils.core import setup

import py2exe

py2exe_options = {"compressed": 1,  # 压缩
                  "optimize": 2,
                  "bundle_files": 1,  # 所有文件打包成一个exe文件
                  # "dll_excludes": ["MSVCP90.dll"]
                  }

setup(console=[{"script": 'report.py',
                "icon_resources": [(1, "templates/rd.ico")]},
               {"script": "init_data.py",
                "icon_resources": [(1, "templates/db_48X48.ico")]}],
      data_files=[("data", ["data/db.xlsx",
                            "data/info.db"]),
                  ("reports", ["reports/list.xlsx"]),
                  ("templates", ["templates/pzb.PNG",
                                 "templates/signature.gif"]),
                  ("templates", glob.glob("templates/*.doc"))],
      name='Report Generator',
      description="Qc Report Generator, RD Software",
      version='1.0.1',
      zipfile=None,  # 不生成library.zip文件
      options={'py2exe': py2exe_options}
      )
