# TeXcel
This program reads tables from an Excel file (.xlsx or xls files) and outputs them in LaTeX format. 

Use the command line interface to use the program.

Use texify -p to transform an Excel table into a LaTeX code. You can customize the output by adding options. Note that the program has been tested only on Windows 10 and 11 and that, in any case, the author doesn't take any responsibility for eventual damages caused to hardware or data by this software. 

This script relies on two packages you may install before using it: 
    pandas - to convert excel tables into dataframes    
    tkinter - to display GUIs



Here follow all the commands and their options.

-copyright: displays the copyright.

-help: shows the help.

-quit: exits the program.

-texify: transforms a sheet into a given Excel file into a LaTeX code.
Options:
    -p specifies the path of the file. If no argument is specified a window will be prompted.
    -s specifies the sheet name or number. Use an integer for the number or a string (e.g. "my data") for the name. You may add multiple numbers by separating them          with blanks. 
    -h specifies the row where to start (the header containin column names)
    -c specifies the columns to be used. Can be integer, string or a list. If you want to select separate columns use "1,2,4" or "A,B,D"; if you want to specify an interval use the form "A:D".
    -n specifies a list of names to be used as header. Separate the names with a space (e.g. texify -n name age date;)
    -T specifies the title of the table in LaTeX
    -L specifies the label to use in LaTeX
    -D specifies the divisors of the table to use (e.g. {l|c|r})
    -o specifies to save the output in a file. If no path is provided a window will be prompted

-setwd: changes the working directory of the script. A window will prompted.
Options:
    -m if this option is provided a manual path must be specified after it.
