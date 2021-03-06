# TeXcel


## Installing Texcel 


### What is Texcel
This program reads tables from an Excel file (.xlsx or xls. files) and outputs them in LaTeX format. 

To install TeXcel you can either use **pip** or download this branch.

This script relies on three packages: 

   -  pandas - to convert excel tables into dataframes    
    
   -  tkinter - to display GUIs
    
   -  openpyxl - necessary to interact with excel files
    



### Installing with PIP


To install TeXcel through PIP use the following command.

On Windows:

      py -m pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple texcel

On MacOS or on Linux:

      python3 -m pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple texcel


**Note that** for the moment TeXcel is on the testing server. You may want to use a [virtual enviroment](https://packaging.python.org/en/latest/guides/installing-using-pip-and-virtual-environments/)  to use the program. 




### Installing from this source


To install this program directly from source you can download this folder, open the terminal (e.g. cmd.exe on Windows), set the working directory on this folder:

      cd \the\path\of\this\folder\texcel-main

Build the package using wheel: 

On Windows

      py -m setup.py sdist bdist_wheel

On MacOS and Linux:

      python3 -m setup.py sdist bdist_wheel


Then you can install the package:

On Windows 

      py -m pip install -e .

On MacOS and Linux:

      python3 -m pip install -e .



## Where to start


To use TeXcel you must have Python 3 installed on you computer. First of all, start Python on your terminal:

Type <code>py</code> on Windows or <code>python3</code> on MacOS and Linux .

Then just import TeXcel and the program will open.

      >>> import texcel

Once the program is running, type 

    texify -p
    
A window to choose a file to read will be prompted to you, and the result in LaTeX will be displayed directly in the terminal. If you want to save the output in a file (expecially if it is a long one) you may write

    texify -p -o
    
_Note that_ if you specify a path after -p or -o, the program will use that instead of prompting a window. For instance

    texify -p myfile.xlsx -o myoutput.txt

_If you don't specify any other option_, the first sheet of the file will selected. The header will be the first row and all the columns with at least one data will be read. 
The following specification 
   
    texify -p -o -s 1 -h 5 -c B:D -n name surname age -L tab1
    
will read the second sheet in the file, using as header row 6 (Python starts from 0), as columns B, C and D and customizing the header ad "name", "surname" and "age". Also, the label in LaTeX of the table will be tab1. 

If you need a _long table_ (e.g. *a multipage table*) you can use the package "longtable" in LaTeX. To convert a table in the correct format, you can use, intead of texify, the command "longtable"

      longtable -p -o 



### Formatting columns


Another issue of Excel is that even if a column of data is formatted to have a certain number of decimal places, or to display a currency, the _real_ data in Excel are still integers or floats (only if decimal places are different from 0). You can tell TeXcel to format a column to place a symbol before the numbers and to display a certain amount of decimal places. 
For instance, if column 2 of you database contains euro with 2 decimal places, you can use the option 

      -f ???.2 1
   
where the first string is "???.2" and tells the program to use the euro symbol and to diplay 2 places, and the second one is the number of the column (remember that Python starts from 0, so the second column is column number 1).
Formatting rules must contain the dot even if one of the two parts is useless. For instance, in order to use only the euro symbol you shall use "???." as rule, and to display only the 2 digits you shall use ".2". You can concatenate as many rules as you want. 

If you use a symbol the char "%", TeXcel will automatically format the string in order to display a percentage. So 

    -f %.2 1
  
won't output something like "%50.22" but instead "50.22%". 


### Restarting the program

If you closed TeXcel but you want to re-use in the same instance of Python where you imported it, use the function start():

      >>> texcel.start()

This will reopen the console and let you use the program. 



## Commands and options


 -  **copyright**: displays the copyright.

 -  **help**: shows the help.

 -  **quit**: exits the program.

 -  **texify**: transforms a sheet into a given Excel file into a LaTeX code. Options for texify:
      -  -p specifies the path of the file. If no argument is specified a window will be prompted.
    
      -  -s specifies the sheet name or number. Use an integer for the number or a string (e.g. "my data") for the name. You may add multiple numbers by separating them          with blanks. 
    
      -  -h  specifies the row where to start (the header containin column names)
    
      -  -c  specifies the columns to be used. Can be integer, string or a list. If you want to select separate columns use 1,2,4 or A,B,D; if you want to specify an interval use the form A:D. **Never put a space between the letters** (e.g. _don't_ write A, B, C).
    
      -  -n  specifies a list of names to be used as header. Separate the names with a space (e.g. texify -n name age date;)
   
      -  -f  formats specific columns of the table adding a symbol before the value and choosing how many decimal digits to show. Use $.2 to use the dollar symbol and show two decimal digits. You can also specify only one of the two elements, using for instace .2 or $. ; the dot must be present in anycase.
    
      -  -T  specifies the title of the table in LaTeX
    
      -  -L  specifies the label to use in LaTeX
    
      -  -D  specifies the divisors of the table to use (e.g. {l|c|r})
   
      -  -R  tells the program to add an horizontal line for each row
    
      -  -o  specifies to save the output in a file. If no path is provided a window will be prompted
    
    

 -  **longtable**: works exaclty like texify, but outputs the table using the LaTeX package "longtable" format; it's useful when the table contains many rows and has to be displayed on multiple pages.  
Options for longtable: **the same of texify**


 -  **setwd**: changes the working directory of the script. A window will prompted. Options for setwd:

      -  -m if this option is provided a manual path must be specified after it.
