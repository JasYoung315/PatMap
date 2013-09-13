from distutils.core import setup
import py2exe
import sys
from glob import glob

data_files = [("Microsoft.VC90.CRT", glob(r'C:\Program Files\Microsoft Visual Studio 9.0\VC\redist\x86\Microsoft.VC90.CRT\*.dll'))]
setup(
    data_files=data_files,
    options = sys.path.append("C:\\Program Files\\Microsoft Visual Studio 9.0\\VC\\redist\\x86\\Microsoft.VC90.CRT"),
)

setup(console = ["Main.py"])
