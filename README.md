#TCC_Report_Create
##### Dominic Cupo

## Usage
There are currently two methods to use for this program.

### Method #1 - `default` flag. 

Provide the program with the `--default` flag. The program will look for files in the predetermined file structure. 

```
Working_Folder/ (call script from here)
    PDF/
        Coordination*.pdf
    TCCs/
        TCC_Base_v*.pdf
        TCC_Rec_v*.pdf
```

Note that the `Working_Folder` is not necessarily where the python script is, just where it is called from

### Method #2 - Individual file hooks
If using this method, do not give the `--default flag`, instead, you will individually point the program at the three files. 

The commands are:

`--cord_path` - The path to the coordination file

`--base_path` - The path to the Base TCC file

`--rec_path` - The path to the Recommended TCC File

If the file has a space in the name, wrap the entire path to the file in `" "`, e.g. :

`--cord_path "C:/path/to/file with space.pdf"`

You can use `/` instead of the normal windows separator `\ ` without issue.

### Note: Coordination file name
The Coordination file can be flagged with a 'CE' or 'RH' in the file name. Whichever name the file has, will be matched in the output file name. 