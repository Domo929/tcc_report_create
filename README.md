#TCC_Report_Create
##### Dominic Cupo

## Installation
Navigate to the [Python Downloads Page](https://www.python.org/downloads/) 

Install the latest `3.7.x` version. Make sure `Add Python to PATH` is selected.
Leave the default install location alone, Python has issues if you install it in other locations you don't own. 

Open up the commandline (`cmd`) and run `python --version`. It should return `3.7.x`

Still in commandline, navigate to the root of the `TCC_Create_Report` project. You can use `cd` to choose the directory, 
and `dir` to list out what directories and files are in your current directory. Verify that the `dir` call in `TCC_Create_Report` 
contains `requirements.txt`

Then, run `pip3 install -r requirements.txt`. This should return no errors.

From here, you can run the script as normal

## Usage
There are currently two methods to use for this program.

### Method #0 - Using file prompts

If you just run the `TCC_Report_Create.py` file without any commandline arguments, it will prompt you for the mode and files.

### Method #1 - `default` flag. 

Provide the program with the `--default` (`-d`) flag. The program will look for files in the predetermined file structure. 

```
Working_Folder/ (call script from here)
    PDF/
        8.0 - Coordination Results & Recommendations_*2018.pdf
    TCCs/
        TCC_Base_v*.pdf
        TCC_Rec_v*.pdf
```

Note that the `Working_Folder` is not necessarily where the python script is, just where it is called from

### Method #2 - Individual file hooks
If using this method, do not give the `--default flag`, instead, you will individually point the program at the three files. 

The commands are:

`--cord_path` (`-c`) - The path to the coordination file

`--base_path` (`-b`) - The path to the Base TCC file

`--rec_path` (`-r`) - The path to the Recommended TCC File

If the file has a space in the name, wrap the entire path to the file in `" "`, e.g. :

`--cord_path "C:/path/to/file with space.pdf"`

You can use `/` instead of the normal windows separator `\ ` without issue.

## Matching

If you want to zip the files together by checking and comparing TCC names you have to specify the `--matching` (`-m`) 
command. This will take significantly longer, but should fix issues with out-of-order files. 

## Notes
### Matching
For the matching to work correctly, TCC names must match EXACTLY, down to the `_` and `-` used. Please verify they match 
between scenarios.   

### Coordination file name
The Coordination file can be flagged with a 'CE' or 'RH' in the file name. Whichever name the file has, will be matched
 in the output file name. 

### Logging
The log of the most recent run will be saved to the `TCC_Create_Logs.log` 
file in the same directory as `TCC_Report_Create.py`.