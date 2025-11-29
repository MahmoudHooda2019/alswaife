import PyInstaller.__main__
import os
import shutil

# Clean previous build
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')

# Define resources to include
# Format: "source;dest" for Windows
res_path = os.path.join("res")
add_data = f"{res_path};res"

PyInstaller.__main__.run([
    'main.py',
    '--name=AlSawifeFactory',
    '--onefile',
    '--noconsole',
    f'--add-data={add_data}',
    '--icon=res/icon.ico',
    '--clean',
])
