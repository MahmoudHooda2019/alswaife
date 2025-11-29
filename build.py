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
version_file = "version.py"
changelog_file = "CHANGELOG.md"
readme_file = "README.md"
release_notes_file = "RELEASE_NOTES.md"

args = [
    'main.py',
    '--name=AlSawifeFactory',
    '--onefile',
    '--noconsole',
    f'--add-data={res_path};res',
    f'--add-data={version_file};.',
    f'--add-data={changelog_file};.',
    f'--add-data={readme_file};.',
    f'--add-data={release_notes_file};.',
    '--icon=res/icon.ico',
    '--clean',
]

PyInstaller.__main__.run(args)
