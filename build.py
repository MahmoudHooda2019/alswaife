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

# Create README.md if it doesn't exist
if not os.path.exists(readme_file):
    with open(readme_file, 'w', encoding='utf-8') as f:
        f.write("# Al Sawife Factory Management System\n\n")
        f.write("A comprehensive factory management system for Al Sawife Factory.\n\n")
        f.write("## Features\n")
        f.write("- Invoice Management\n")
        f.write("- Attendance & Departure Tracking\n")
        f.write("- Employee Management\n")
        f.write("- Excel Export Capabilities\n\n")
        f.write("## Installation\n")
        f.write("Run the executable installer or extract the portable version.\n\n")
        f.write("## Usage\n")
        f.write("Launch AlSawifeFactory.exe to start the application.\n")

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