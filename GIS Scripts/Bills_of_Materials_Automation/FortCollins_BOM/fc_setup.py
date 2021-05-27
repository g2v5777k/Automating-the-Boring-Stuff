from distutils.core import setup
# from glob import glob
import py2exe
import matplotlib

# directory = r'C:\Program Files (x86)\Common Files\ArcGIS\bin\Microsoft.VC90.CRT\*.*'
# directory=directory.rstrip()
dll_excludes = ['MSVCP90.dll',
                'api-ms-win-crt-runtime-l1-1-0.dll',
                'python37.dll','api-ms-win-crt-environment-l1-1-0.dll',
                'api-ms-win-crt-filesystem-l1-1-0.dll',
                'api-ms-win-crt-heap-l1-1-0.dll',
                'api-ms-win-crt-convert-l1-1-0.dll',
                'api-ms-win-crt-string-l1-1-0.dll',
                'api-ms-win-crt-math-l1-1-0.dll',
                'api-ms-win-crt-utility-l1-1-0.dll',
                'api-ms-win-crt-stdio-l1-1-0.dll',
                'api-ms-win-crt-time-l1-1-0.dll',
                'OLEAUT32.dll',
                'IMM32.dll'
                'SHELL32.dll'
                'COMDLG32.dll',
                'COMCTL32.dll'
                'WS2_32.dll',
                'ole32.dll',
                'MSVCP90.dll',
                'IPHLPAPI.DLL',
                'NSI.dll',
                'WINNSI.DLL',
                'WTSAPI32.dll',
                'SHFOLDER.dll',
                'PSAPI.dll',
                'MSVCR120.dll',
                'MSVCP120.dll',
                'CRYPT32.dll',
                'GDI32.dll',
                'ADVAPI32.dll',
                'CFGMGR32.dll',
                'USER32.dll',
                'POWRPROF.dll',
                'MSIMG32.dll',
                'WINSTA.dll',
                'MSVCR90.dll',
                'KERNEL32.dll',
                'MPR.dll',
                'Secur32.dll',
                'WS2_32.dll',
                'COMCTL32.dll',
                'IMM32.dll',
                'SHELL32.dll',
                'COMDLG32.dll'
                ]

options = {"py2exe": {"excludes":["arcpy"],
                     "dll_excludes":dll_excludes,
                     "includes":["lxml._elementpath", "numpy", "xml.etree.ElementTree", "xlwt", 'UserList', 'UserString'],
                     "packages": ["email", "uuid", "lxml"]}}
setup(
      console=[r'C:\Users\ColtonLuttrell\Desktop\Desktop_Files\Python\Market_Specific_Scripts\FC_BOM\FC_BOM_v11.py'],
      data_files=matplotlib.get_py2exe_datafiles(),
      options=options
      )
