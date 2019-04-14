from cx_Freeze import setup, Executable

setup(
    name = "ADdUser",
    version = "0.1",
    description = "This program adds user from a excel file to an AD server",
    executables = [Executable("main.py")],
)
