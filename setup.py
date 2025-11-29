from setuptools import setup, find_packages
import os

# Read the contents of requirements.txt
def read_requirements():
    with open('requirements.txt') as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#')]

# Read the contents of README.md
def read_readme():
    with open('README.md', encoding='utf-8') as f:
        return f.read()

# Get version from version.py
def get_version():
    version = {}
    with open(os.path.join('version.py')) as f:
        exec(f.read(), version)
    return version['__version__']

setup(
    name='al-sawife-factory',
    version=get_version(),
    description='A GUI application for managing factory operations and invoices',
    long_description=read_readme(),
    long_description_content_type='text/markdown',
    author='AlSawife Factory',
    author_email='mh20192004@gmail.com',
    url='https://github.com/MahmoudHooda2019/alsawife',
    packages=find_packages(),
    include_package_data=True,
    install_requires=read_requirements(),
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Manufacturing',
        'Topic :: Office/Business :: Financial :: Accounting',
        'License :: Other/Proprietary License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
    ],
    python_requires='>=3.8',
    entry_points={
        'console_scripts': [
            'alsawife-factory=main:main',
        ],
    },
    package_data={
        '': ['res/*.json', 'res/*.ico'],
    },
    zip_safe=False,
)