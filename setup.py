from cx_Freeze import setup, Executable
  
setup(name = "ExcelProject" ,
      version = "0.1" ,
      options = {"build_exe":{"packages":["xlrd","xlwt","xlutils"]}},
      description = "Excel Automation Project" ,
      executables = [Executable("PythonProject.py")])

