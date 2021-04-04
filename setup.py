from cx_Freeze import setup, Executable

base = "Win32GUI"

includefiles = ['config.ini']
executables = [Executable("KLog.pyw", icon="C:\\Users\\Kenneth\\PycharmProjects\\KLog\\logo.ico", base=base)]

packages = ["idna"]
options = {
    'build_exe': {
        'packages':packages,
        'include_files':includefiles,
        'excludes':['tkinter', 'numpy'],
        'optimize': 2
    },
}

setup(
    name = "KLog",
    options = options,
    version = "1.0",
    description = 'KLog',
    executables = executables, requires=['pyautogui', 'pynput']
)