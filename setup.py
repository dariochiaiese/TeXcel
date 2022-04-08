from setuptools import setup

setup(name='texcel',
      version='3.2.3',
      description='This program reads Excel tables from .xlsx (or .xls) files and outputs them in LaTeX fromat',
      url='http://github.com/dariochiaiese/texcel',
      author='Dario Chiaiese',
      author_email='dariochiaiese@gmail.com',
      license='GPLv3',
      packages=['texcel'],
      install_requires=[
          'Tk==0.1.0',
          'pandas',
          'openpyxl',
      ],
      include_package_data=True,
      package_data={'': ['*.txt']},
      zip_safe=False)
