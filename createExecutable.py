from cx_Freeze import setup, Executable

setup(name = "WordToPdf" ,
      version = "0.1" ,
      description = "" ,
      executables = [Executable("wordtopdf.py")])