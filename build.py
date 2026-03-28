"""PyInstaller 빌드 스크립트 — Windows에서 실행"""
import PyInstaller.__main__

PyInstaller.__main__.run([
    'app.py',
    '--onefile',
    '--windowed',
    '--name=SKU_변환기',
    '--add-data=core;core',
    '--hidden-import=openpyxl',
    '--hidden-import=pandas',
    '--hidden-import=rapidfuzz',
])
