from setuptools import setup, find_packages

APP = ['classgui.py', ]

DATAFILES = [('assets', ['assets/logo.jpg', 'assets/icon.icns'])]
OPTIONS = {
           'packages': ['tkinter', 'openpyxl', 'PIL'],
           'resources': DATAFILES,
           'iconfile':'icon.icns',
           'plist':{
               'CFBundleName': 'HybridQC',
               'CFBundleDevelopmentRegion': 'English',
               'CFBundleIdentifier': 'com.iita_ayatoo.classgui',
               'CFBundleVersion': '1.0.0',
           }}

setup(
    app=APP,
    data_files=DATAFILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app']
)
