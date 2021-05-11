from distutils.core import setup
# from glob import glob
import py2exe
import matplotlib

# directory = r'C:\Program Files (x86)\Common Files\ArcGIS\bin\Microsoft.VC90.CRT\*.*'
# directory=directory.rstrip()
dll_excludes = ['MSVCP90.dll']

options = {"py2exe": {"excludes":["arcpy"],
                     "dll_excludes":dll_excludes,
                     "includes":["lxml._elementpath", "numpy", "xml.etree.ElementTree", "xlwt", 'UserList', 'UserString'],
                     "packages": ["email", "uuid", "lxml"]}}
setup(
      console=[r'C:\Users\ColtonLuttrell\Desktop\Desktop_Files\Python\Market_Specific_Scripts\Clarity_BOMs\Clarity_BOM_GUI37.py'],
      data_files=matplotlib.get_py2exe_datafiles(),
      options=options
      )
