'''
Program: TeXcel
Author: Dario Chiaiese
Version: 3.2.3
Licence: GPLv3

Description: This program connects to an Excel file, abstracts a table and finally outputs it in LaTeX format. 

Dependencies: 
    pandas - to convert excel tables into dataframes    
    tkinter - to display GUIs
    openpyxl - to interact with excel files
'''

import os
from os import read
import pandas
from tkinter import Tk
from tkinter.filedialog import askdirectory, askopenfilename, asksaveasfilename #we'll need it in order to let the user choose the excel file he wants to convert
import openpyxl


#-----------------------------------------------------------------ENVIRONMENT VARIABLES---------------------------------------------------

#Special symbols for the column formatting function
SPEC_UNITS = ["%","kg","g","lbs","m","km","cm","mm","nm"]

#------------------------------------------------------------------------MAIN------------------------------------------------------------------

def main():
    print("""Welcome to TeXcel! Type help; to display the help of the program. 
    Use texify -p to convert a file and add options to customize the output.
    Type licence to display the full licence of the program.""")   
    
    path = os.path.dirname(os.path.realpath(__file__))
    os.chdir(path)
    
    readfile("copyright.txt")    
    console()

#-------------------------------------------------------------------PANDAS FUNCTIONS------------------------------------------------------------------


def read_exc(path, sn = [0], hd= 0, nms=None, cols = None):
    #cols indicates which columns to read. If None, all columns are used. If str, it can be "B:D" to indicate a range, "A,E" to indicate separate 
    # columns or even "A, C:F" to mix; if list of int it can indicate separate colums; the same if list of strings.
    if not path.replace(" ", ""):
        return print("The operation could not be completed as the user did not give a valid path.")
    
    try:
        data = pandas.read_excel(io = path, sheet_name = sn, header= hd, names=nms, usecols = cols)
        mats = []
        for sheet in data.values(): #THE INITIAL data IS A DICTIONARY WHERE THE KEYS ARE THE SHEET NUMBER, EVEN IF THERE IS ONLY ONE SHEET
            mats.append(create_matrix(sheet)) #each sheet is a pandas dataframe object; converts every pandas dataframe in a matrix
    except Exception as e:
        print("An error was raised during the reading of the Excel file. Error name: ", e)
        return False

    return mats   #mats will be a matrix of matrices; each matrix represents the valuable data of a sheet.


def create_matrix(data):
    #creates the matrix starting from a dataframe object in pandas
    idx = list(data.index)
    header = list(data.columns)
    mat = [header]

    for i in range(len(idx)): #PLEASE note that indexes only concern rows which are not the header. The header has not an index number
        mat.append(list(data.loc[i]))
    
    return mat
   
#----------------------------------------------------------------LATEX FUNCTIONS-------------------------------------------------------------------------
    
def to_latex(mat, title = None, label = None, div = None, divide_row = False):
#takes as input a matrix and transforms it according to the latex format and the options passed by the user
#PLEASE NOTE_ that I will use [] instead of {}, and then replace it, since {} is ambiguous for Python in string formatting
#div is the divisors and the aligment the user wants to use for the table: e.g. {l|c|r}
#divide_row, if true, adds an \hline for each row.
    
    if not div: #If div has not been specified by the user, it must generated automatically
        div = "[|"
        for i in range(len(mat[0])): #there must be as indicator as the columns in the matrix
            div += "l|"
        div += "]"


    latex = """
    \\begin[table](h!)
    \\begin[center]
        \\caption[{}]
        \\label[tab:{}]            
            \\begin[tabular]{}
                \\hline
    \n""".format(title, label, div)  #REMEMBER that "\" is an escape char in python
    
    #creates the header
    header = ""        
    for col in mat[0]: #mat[0] is the header
        header += str(col) + " & " 
    header = header[:-2] + "\\\ \n \\hline \n" #replaces the last, redundand & with the newline command \\
    latex += header
    
    #adds the lines
    for line in mat[1:]:
        row = ""
        for element in line:
            row += str(element) + " & "
        row = row[:-2] + "\\\ \n" + (" \\hline \n " if divide_row else "") 
        latex += row
        
    latex += """
            \\end[tabular]
        \\end[center]
    \\end[table]

    \n ----------------------------------END TABLE-----------------------------------------------
    """
    
    latex = latex.replace("[", "{")
    latex = latex.replace("]", "}")
    latex = latex.replace("(","[")
    latex = latex.replace(")","]")
    
    return latex


def to_latex_longtable(mat, title = None, label = None, div = None, divide_row = False):
#takes as input a matrix and transforms it according to the latex package "longtable"  (user should import package in Latex by
# \usepackage{longtable})
#PLEASE NOTE_ that I will use [] instead of {}, and then replace it, since {} is ambiguous for Python in string formatting
#div is the divisors and the aligment the user wants to use for the table: e.g. {l|c|r}
#divide_row, if true, adds an \hline for each row.
    
    if not div: #If div has not been specified by the user, it must generated automatically
        div = "[|"
        for i in range(len(mat[0])): #there must be as indicator as the columns in the matrix
            div += "l|"
        div += "]"


    latex = """
    \\begin[longtable](h!){}    
        \\caption[{}\label[{}]] \\\          
            \\hline
            \\multicolumn[{}][| c |][Content of the table]
            \\hline               
    \n""".format(div, title, label, len(mat[0]))  #REMEMBER that "\" is an escape char in python
    
    #creates the header
    header = ""        
    for col in mat[0]: #mat[0] is the header
        header += str(col) + " & " 
    header = header[:-2] + "\\\ \n \\hline \n" #replaces the last, redundand & with the newline command \\
        
    latex += header #header in the first head
    latex += """
            \\endfirsthead
            \\hline
            \n"""
    latex += header #header repeated in every new page
    latex += """
            \\endhead
            
            \\hline
            \\endfoot               
    \n"""
    
    
    #adds the lines
    for line in mat[1:]:
        row = ""
        for element in line:
            row += str(element) + " & "
        row = row[:-2] + "\\\ \n" + (" \\hline \n " if divide_row else "") 
        latex += row
        
    latex += """            
        \\end[longtable]

    \n ----------------------------------END TABLE-----------------------------------------------
    """
    
    latex = latex.replace("[", "{")
    latex = latex.replace("]", "}")
    latex = latex.replace("(","[")
    latex = latex.replace(")","]")
    
    return latex


def format_column(rules, mat):
    #rules is a list of couples like [ [form1, col1], [form2, col2], ...  ]
    #tranforms mats columns using the formatting symbol.digits . User must use the dot even if symbol or digits is useless
    # e.g. "$." ".2" are both legal formattings
    
    for rule in rules:
        
        col = int(rule[1]) #column to be modified
        if not "." in rule[0]: raise Exception("Formatting rules must be formed as follows: sym.dec sym. or .dec")
        rule = rule[0].split(".")
        sym = rule[0] #the symbol to be placed before the digits
        dec = rule[1] #the number of decimal places
        if not dec or not str(dec).isdigit() or int(dec) < 0: 
            dec = 0
        else:
            dec = int(dec) #decimal positions to apply        
        
        for row in mat[1:]: #header must be excluded
            if not str(row[col]).replace(".","").isdigit(): raise Exception("Column {} does not cointain only numbers!".format(col)) #if the column does not contain numbers, then an error must be raised
            if sym.lower() in SPEC_UNITS: #special chars are to be placed after the digits
                row[col] = "%.{}f".format(dec) % row[col] + sym #e.g. if dec = 2 and sym="%" transforms 12 into 12.00%
            else:
                row[col] = sym + "%.{}f".format(dec) % row[col] #e.g. if dec = 2 and sym="$" transforms 1230 into $1230.00
        
        
    return mat 
            
 
#----------------------------------------------------------------OUTPUT-------------------------------------------------------------------------        
    
def print_output(path, mats):
#prints the output in latex into a file chosen by the user
    try:            
        f = open(path, 'a+')
        for mat in mats: #now mats is a list of strings representing each table in latex
            f.write(mat)
    finally:
        f.close()

def readfile(path):
    #reads a file
    try:
        f = open(path, 'r', encoding="utf8")
        print(f.read())
    except Exception as e:
        print("An error occured reading the file. The error is :", e)
    finally:
        f.close()
    
#----------------------------------------------------------------CONSOLE-------------------------------------------------------------------------

def console():
    #Console command must have the form command, -o1 arg1, -o2 arg2, ...
    args = input("TeXcel console ~ ")
    args = args.strip() + ";" #adds ; at the end
    print(args)
    args = command_breaker(args)
    launch_console(args)


def launch_console(args):
    '''
    args contains all the arguments to be used in the console.
    args is a list of strings in the form ["command", ["-o1", "val1"], ["-o2", "val2"], ...] 
    -p specifies the path of the file
    -s specifies the sheet name or number
    -h specifies the row where to start (the header)
    -c specifies the columns to be used. Can be integer, string or a list
    -n specifies a list of names to be used as header
    -f specifies a formatting style for the column. -f for1 col1 for 2 col2 ...
        E.g. $.2 indicates to place $ sign before the number and impose two decimal positions. 
        User can write -f $.2 4 .2 6 to indicate that column 4 must have the first formatting and columnd 6 must have the other
    -T specifies the title of the tabel
    -L specifies the label to use
    -D specifies the divisors to use (e.g. {l|c|r})
    -R if called, places an \hline for each line
    -o specifies to save the output in a file
    '''

    cmd = args[0] #the first element of args is the command to be executed    
    args = args[1:] #cuts off the main command
    print("command is", cmd, "args are", args)

    if cmd == "texify" or cmd == "longtable": 
        opt = {"path": None, 
        "sheet_name": [0], 
        "header": 0, 
        "names": None, 
        "usecols": None,
        "title" : "",
        "label" : "",
        "divisors" : "",
        "divide_row" : False,
        "formatting" : [],
        "output" : "",
        "err" : ""} #a dictionary containing every argument

        for i in args:           
            r = read_texify(i)
            opt[r[0]] = r[1]
        
        if  opt["err"]:
            print("An invalid command was called. The message is " + opt["err"])
            return console()

        print("Opt are ", opt)

        if not opt["path"]:
            opt["path"] = read_texify(["-p"])[1]
        mats = read_exc(opt["path"], opt["sheet_name"], opt["header"], opt["names"], opt["usecols"])
        if not mats: return console() #if something goes wrong, read_exc return None
        
        if opt["formatting"]: #checks if formatting is imposed by the user
            try:
                for mat in mats:
                    mat = format_column(opt["formatting"], mat)
            except Exception as e:
                print("Formatting was invalid! Error is ", e)
            
            
        latextables = []
        
        if cmd == "texify":
            for mat in mats:
                latextables.append(to_latex(mat, opt["title"], opt["label"], opt["divisors"], opt["divide_row"]))
        
        elif cmd == "longtable":
            for mat in mats:
                latextables.append(to_latex_longtable(mat, opt["title"], opt["label"], opt["divisors"], opt["divide_row"]))
           
        
        if opt["output"]: #saves the tables into a file
            print_output(opt["output"], latextables)
        else:
            for table in latextables: #just prints the tables
                print(table)
                
    
    elif cmd == "setwd": #sets the working directory to a new path. If the user doesn not specify anything, a dialog opens
        path = ""
        if args and args[0][0] == "-m": #m stands for manual input and must be followed by a valid path without spaces
            path = args[0][1]
            if not os.path.isdir(path): return print(path, " is not an existing directory directory")
        else:
            try:
                Tk().withdraw()
                path = askdirectory()
            except: return print("The selected directory raised an error. Choose a different one.")
        set_working_directory(path)
        print("New working directoy set to ", os.getcwd())

    elif cmd == "help":
        readfile("help.txt")

    elif cmd == "copyright":
        readfile("copyright.txt")

    elif cmd == "test": 
        print("This is a test mode. Other arguments are", args) 

    elif cmd == "error":
        print("An error occured. The error is: ", args[0])   

    elif cmd == "quit": 
        print("Thanks for using the program! Bye!")
        return input()

    else:
        print("This is not a known command")

    return console()
  

def read_texify(opts):
    #read a couple of console commands (option + value) and returns them
    #opts must be a list of two elements. First one is the option, second one is value
    console_dict = {
        "-p":"path",
        "-s":"sheet_name",
        "-h":"header",
        "-c":"usecols",
        "-n":"names",
        "-e":"err",
        "-f":"formatting",
        "-L":"label",
        "-T":"title",
        "-D":"divisors",
        "-R":"divide_row",
        "-o": "output"
    }
    
    if not opts[0] or not opts[0] in console_dict:
        return ["err", "One or more options were not valid"]

    if opts[0] == "-p" and (len(opts) == 1 or not opts[1]): #GUI option for choosing the path. If -p is alone a prompt to the user is shown
        o = console_dict[opts[0]]
        v = open_dialog("open")
        return [o,v]

    if opts[0] == "-o" and (len(opts) == 1  or not opts[1]): #GUI option for choosing where to save the output
        o = console_dict[opts[0]]
        v = open_dialog("save")
        return [o,v]

    if opts[0] == "-R" and (len(opts) == 1  or not opts[1]): #option -R needs to be alone
        return [console_dict[opts[0]], True]

    if opts[0] and not opts[1]: #from now on all the options need an argument to be valid
        return ["err", "No specified argument for option {}".format(opts[0])]

    if opts[0] == "-s":
    #casts the sheet name into an integer if an integer is passed as argument; otherwise keeps the string with the sheet name
        return ["sheet_name", [ int(opts[1]) if opts[1].isdigit() else opts[1] ]]  

    if opts[0] == "-n":
    #names must be an array of custom names for the columns
        if opts[-1] == "": opts = opts[:-1] #last element cannot be the empty space, otherwise the columns do not match!
        return [console_dict[opts[0]], opts[1:]]
    
    if opts[0] == "-h" and str(opts[1]).isdigit():
        opts[1] = int(opts[1]) #the header must be a number and not a string
        
    if opts[0] == "-f":
        couples = []
        if len(opts[1:])%2 == 0: #the optional arguments must be even (couples of format, column)
            for i in range(1, len(opts[1:]), 2):
                couples.append( [opts[i], opts[i+1]] )
        else:
            return ["err" , "The arguments for -f were invalid. Aguments must be '-f format1 column1 format2 column2' and so on"]
        return [console_dict[opts[0]], couples]
                
        
    o = console_dict[opts[0]]
    v = opts[1]
    return [o,v]

def open_dialog(opt):
    #opens a dialog window with tkinter in order to let the use choose his path
    Tk().withdraw() #avoids tkinter to open an empty windows
    try:

        if opt == "open":
            return askopenfilename() #the value of the function will be the chosen path
        elif opt == "save":
            return asksaveasfilename(defaultextension='.txt', filetypes=[("Text file", '*.txt')],
                    title="Create or choose the file where to append the TeXcel output")

    except: return print("The selected directory raised an error. Choose a different one.")


def command_breaker(comm):
    #takes a command (sting) in input and breaks it into a list of sub commands with the main command and the couples -option value

    comm = comm.strip() #trims out blanks at the end and at the beginning    
    if not comm[-1] == ";": return ["error", "Invalid syntax. Final ';' missing"]

    maincmd = comm.split(" ")[0]
    comm = comm[len(maincmd)+1:]  #deletes the main command 

    if not comm: return [maincmd[:-1]] #if comm is empty, then there was only one command in the form command;
        
    coms = [maincmd] #coms will be a matrix containing the main command and lists of -option value couples

    i = 1
    start = 0
    
    while i < len(comm):
        if comm[i] in ["-",";"]:
            comm = comm.strip()
            substr = comm[start+2:i].strip() #this is the string from the first char after option -o to the last char before next option            
            if substr and substr[0] in ("'", '"') and substr[-1] in ("'", '"'): #if the third char (after the first two representing the option) is " or ', and if the last is " or ', then the parameter is a phrase
                pair = [comm[start : start + 2] , comm[start + 2 : i].strip()]
                pair[1] = pair[1][1:-1] #Eliminates the quotemarks at the end and at the start of the string                
                coms.append(pair) #if the option is a phrase (e.g. "My Title"), we don't want to separate accordin to the spaces
            else:
                coms.append(comm[start:i].strip().split(" ")) #appends a sub list made of [option, value] by splitting x[0] using the blank. x[0] is supposed to be a string like "-o1 value"; i-1 is due to the fact that every new option -o must be preceeded by a blank
            start = i
        i += 1
    
    
    return coms


def set_working_directory(path):
    os.chdir(path)

