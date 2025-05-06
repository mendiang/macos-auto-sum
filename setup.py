"""
Script setup untuk membuat aplikasi Auto Sum Status Bar menggunakan py2app.
"""
from setuptools import setup

APP_NAME = 'AutoSumApp'
APP = ['auto_sum_statusbar.py'] # File skrip utama Anda
DATA_FILES = [] # Tambahkan file data lain jika ada (misal ikon)
OPTIONS = {
    'argv_emulation': True, # Memungkinkan drop file ke ikon app (tidak relevan di sini)
    'packages': ['rumps', 'pyperclip', 're'], # Sebutkan library pihak ketiga
    'iconfile': 'icons8-sigma-32.icns', # (Opsional) Path ke file ikon .icns
    'plist': {
        'CFBundleName': APP_NAME,
        'CFBundleDisplayName': APP_NAME,
        'CFBundleGetInfoString': "Auto Sum Clipboard Monitor",
        'CFBundleIdentifier': "com.yourdomain.autosumapp", # Ganti dengan ID unik Anda
        'CFBundleVersion': "0.1.0",
        'CFBundleShortVersionString': "0.1.0",
        'NSHumanReadableCopyright': u"Copyright © 2023, Anda", # Ganti nama Anda
        'LSUIElement': True,  # <-- Penting: Membuat aplikasi berjalan tanpa ikon di Dock
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)