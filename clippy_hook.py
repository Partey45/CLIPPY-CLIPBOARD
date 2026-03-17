import os, glob, shutil
from PyInstaller.utils.hooks import collect_data_files, get_package_paths

# Collect ALL PyQt6 WebEngine data files (pak, locales, resources)
datas = collect_data_files('PyQt6', includes=['**/*.pak', '**/*.dat', '**/*.bin', '**/*.so', '**/*.pyd'])
datas += collect_data_files('PyQt6.Qt6', includes=['**'])
