from cx_Freeze import setup, Executable
import os

os.environ['TCL_LIBRARY'] = r'D:\Python_35\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'D:\Python_35\tcl\tk8.6'

setup(
    name="ZAMSHEET",
    version="0.1",
    description="Zam Document",
    options={"build_exe":{"include_files":[r"D:\Python_35\DLLs\tcl86t.dll",r"D:\Python_35\DLLs\tk86t.dll","banner.png"]}},
    executables= [Executable("zamtest.py", base="Win32GUI")]
    )
