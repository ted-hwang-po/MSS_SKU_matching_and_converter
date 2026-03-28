"""PyInstaller hook: tkinterdnd2의 tkdnd DLL 파일을 번들에 포함"""
from PyInstaller.utils.hooks import collect_data_files
datas = collect_data_files('tkinterdnd2')
