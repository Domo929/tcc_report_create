#!/usr/bin/env python3
# ######### IMPORTS ######### #
from __future__ import print_function

import argparse
import glob
import os
import sys

from PyPDF3 import PdfFileWriter, PdfFileReader


# ######### ERROR HElPER ######### #
def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


# ######### ARGUMENT HANDLER ######### #
parser = argparse.ArgumentParser(description='Create TCC Reports correctly.')
parser.add_argument('--default',
                    action='store_true',
                    help='Set this to find the files via file path, instead of flagging each individually')
parser.add_argument('--cord_path',
                    action='store',
                    help='The path to the Coordination PDF',
                    type=str)
parser.add_argument('--base_path',
                    action='store',
                    help='The path to the Base TCC PDF',
                    type=str)
parser.add_argument('--rec_path',
                    action='store',
                    help='The path to the Recommended TCC PDF',
                    type=str)
# Store the arguments in a dict for easy reference later
opts = vars(parser.parse_args())

# ######### PATH HANDLING ######### #
# Empty paths to be filled depending on the mode chosen (default or otherwise)
cord_path = ''
base_path = ''
rec_path = ''

# If the default flag was not selected
if not opts['default']:
    # Fill the paths. If they weren't provided they will be null
    cord_path = opts['cord_path']
    base_path = opts['base_path']
    rec_path = opts['rec_path']

    # Check that if default was NOT selected, we are provided with the proper file flags
    if not (cord_path and base_path and rec_path):
        eprint('Didn\'t specify default, but also didn\'t give all the individual flags',
               'Either provide all three "--EXMP_path" flags OR the "--default" flag')
        exit()

# If default WAS selected
else:
    # Get the current working directory. This allows the script to be called from wherever,
    # and find the files correctly
    cwd = os.getcwd()

    # This allows us to find files that have different versions
    base_glob = 'TCC-Base_v*.pdf'
    rec_glob = 'TCC-Rec_v*.pdf'
    cord_glob = '*Coordination*.pdf'

    # Build the full path in order to start searching
    cord_path = os.path.join(cwd, 'PDF', cord_glob)
    base_path = os.path.join(cwd, 'TCCs', base_glob)
    rec_path = os.path.join(cwd, 'TCCs', rec_glob)

    # Build the list of all possible matches to the glob
    base_glob_list = glob.glob(base_path)

    # If there is more than one, ask them to remove the extra
    # TODO Load a file browser and ask the user to choose which one
    if len(base_glob_list) > 1:
        eprint('Found multiple versions of TCC_Base. Please remove all extras')
        exit()
    # If there aren't any files in the folder, throw error
    # TODO Load a file browser and ask the user to choose where the file is
    elif len(base_glob_list) == 0:
        eprint('Found no version of TCC_base. Please provide a sheet')
        exit()
    else:
        base_path = base_glob_list[0]

    # The same checks for the base file happen for the reccomended file
    rec_glob_list = glob.glob(rec_path)
    if len(rec_glob_list) > 1:
        eprint('Found multiple versions of TCC_Rec. Please remove all extras')
        exit()
    elif len(rec_glob_list) == 0:
        eprint('Found no version of TCC_rec. Please provide a sheet')
        exit()
    else:
        rec_path = rec_glob_list[0]

    # Finally, the same checks happen for the coordination file
    cord_glob_list = glob.glob(cord_path)
    if len(cord_glob_list) > 1:
        eprint('Found multiple versions of Coordination PDF. Please remove all extras')
        exit()
    elif len(cord_glob_list) == 0:
        eprint('Found no version of Coordination PDF. Please provide a sheet')
        exit()
    else:
        cord_path = cord_glob_list[0]

# ######### PDF Output Name Checking ######### #
# Empty string to hold it all
output_name = ''
# This splits the total file path into the list of drive, directory and at the end the file name
paths = cord_path.split(os.sep)
# Pull out the last entry in the list (the file name)
filename = paths[len(paths) - 1]

# Check to see if either is in the name, save for later
if 'CE' in filename:
    output_name = 'CE'
elif 'RH' in filename:
    output_name = 'RH'

# ######### PDF Write Setup ######### #
# Open the output file writer
output = PdfFileWriter()

# Open the input PDFs
cord_in = PdfFileReader(open(cord_path, 'rb'), False)
base_in = PdfFileReader(open(base_path, 'rb'), False)
rec_in = PdfFileReader(open(rec_path, 'rb'), False)

# Check that the same number of pages exist in the Base and Recommended PDFs.
# They should be the same length
if base_in.getNumPages() != rec_in.getNumPages():
    eprint('Number of pages does not match in the Base and Recommended PDFs')
    exit()

# Check that the coordination PDF is longer than the base (and therefore recc) pdf too.
# The Coordination PDF includes pages at the front that do not get sliced in, and instead actually sit
# in the front. If the Coordination pdf is less than the Base or Recc, these are missing, or there was another error
if cord_in.getNumPages() < base_in.getNumPages():
    eprint('Coordination PDF is shorter than the Base/Recc PDFs')
    exit()

# ######### ZIPPING ######### #
# Find the difference in length of the PDFs, these are the leader pages of the coordination
diff_length = cord_in.getNumPages() - base_in.getNumPages()

# Add the first pages from the coordination PDF to the output PDF
for ii in range(diff_length):
    output.addPage(cord_in.getPage(ii))

# Now we go through and 'zip' the files together. Note that the Coordination PDF has the diff_length
# offset added to it to account for the starting pages. Everything else is just the actual indices
for jj in range(base_in.getNumPages()):
    output.addPage(cord_in.getPage(jj + diff_length))
    output.addPage(base_in.getPage(jj))
    output.addPage(rec_in.getPage(jj))

# Finally, output everything to the PDF
# The output name is chosen based on what the name of the coordination file is
output_name = "8.0 - Coordination Results & Recommendations_" + output_name + "2018.pdf"
with open(output_name, "wb") as w:
    output.write(w)
