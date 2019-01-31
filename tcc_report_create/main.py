#!/usr/bin/env python3
# ######### IMPORTS ######### #
from __future__ import print_function

import argparse
import glob
import os
import sys

from PyPDF3 import PdfFileWriter, PdfFileReader

from tcc_report_create.helpers import *


# ######### ERROR HElPER ######### #
def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


# ######### ARGUMENT HANDLER ######### #
parser = argparse.ArgumentParser(description='Create TCC Reports correctly.')
parser.add_argument('-d', '--default',
                    action='store_true',
                    help='Set this to find the files via file path, instead of flagging each individually')
parser.add_argument('-c', '--cord_path',
                    action='store',
                    help='The path to the Coordination PDF',
                    type=str)
parser.add_argument('-b', '--base_path',
                    action='store',
                    help='The path to the Base TCC PDF',
                    type=str)
parser.add_argument('-r', '--rec_path',
                    action='store',
                    help='The path to the Recommended TCC PDF',
                    type=str)
parser.add_argument('-m', '--matching',
                    action='store_true',
                    help='Whether or not to trigger the TCC name matching method of combining')
# Store the arguments in a dict for easy reference later
opts = vars(parser.parse_args())

# ######### PATH HANDLING ######### #
# Empty paths to be filled depending on the mode chosen (default or otherwise)
cord_path = ''
base_path = ''
rec_path = ''

recc_glob_list = []
recc_pdf_exists = False

# If the default flag was not selected
if not opts['default']:
    # Fill the paths. If they weren't provided they will be null
    cord_path = opts['cord_path']
    base_path = opts['base_path']
    rec_path = opts['rec_path']

    # Check that if default was NOT selected, we are provided with the proper file flags
    if not (cord_path and base_path):
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

    # The same checks for the base file happen for the recommended file
    recc_glob_list = glob.glob(rec_path)
    recc_pdf_exists = len(recc_glob_list) > 0
    if recc_pdf_exists:
        if len(recc_glob_list) > 1:
            eprint('Found multiple versions of TCC_Rec. Please remove all extras')
            exit()
        elif len(recc_glob_list) == 0:
            eprint('Found no version of TCC_rec. Please provide a sheet')
            exit()
        else:
            rec_path = recc_glob_list[0]

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
else:
    output_name = ''

# ######### PDF Write Setup ######### #
# Open the input PDFs
cord_pdf = PdfFileReader(open(cord_path, 'rb'), False)
base_pdf = PdfFileReader(open(base_path, 'rb'), False)
recc_pdf = ''
if recc_pdf_exists:
    recc_pdf = PdfFileReader(open(rec_path, 'rb'), False)

# Check that the same number of pages exist in the Base and Recommended PDFs.
# They should be the same length
if recc_pdf_exists:
    if base_pdf.getNumPages() != recc_pdf.getNumPages():
        eprint('Number of pages does not match in the Base and Recommended PDFs')
        exit()

# Check that the coordination PDF is longer than the base (and therefore recc) pdf too.
# The Coordination PDF includes pages at the front that do not get sliced in, and instead actually sit
# in the front. If the Coordination pdf is less than the Base or Recc, these are missing, or there was another error
if cord_pdf.getNumPages() < base_pdf.getNumPages():
    eprint('Coordination PDF is shorter than the Base/Recc PDFs')
    exit()

# Find the difference in length of the PDFs, these are the leader pages of the coordination
diff_length = cord_pdf.getNumPages() - base_pdf.getNumPages()

# ######### ZIPPING ######### #

# Open the output file writer
output = PdfFileWriter()

if opts['matching']:
    print("Converting Coordination PDF to string")
    cord_str_pages = pdf_pages_to_list_of_strings(cord_path)

    print("Converting Base PDF to string")
    base_str_pages = pdf_pages_to_list_of_strings(base_path)

    recc_str_pages = []
    if recc_pdf_exists:
        print("Converting Recommended PDF to string")
        recc_str_pages = pdf_pages_to_list_of_strings(rec_path)

    # regex_cord = re.compile(r"\bTCC_\*")
    regex_cord = r"(TCC Curve: )(TCC_[\w/ \[\]\"-]+)"
    regex_base_recc = r"(TCC Name: )(TCC_[\w/ \[\]\"-]+)"

    for ii in range(diff_length, len(cord_str_pages)):
        output.addPage(cord_pdf.getPage(ii))
        matches = re.finditer(regex_cord, cord_str_pages[ii], re.MULTILINE)
        for match_num, match in enumerate(matches, start=1):
            recc_num = 0
            if recc_pdf_exists:
                recc_num = find_matching_page(match.group(2), recc_str_pages, regex_base_recc)
            print("Attempting to find: " + match.group(2))
            base_num = find_matching_page(match.group(2), base_str_pages, regex_base_recc)
            print("Found on base page: " + str(base_num))
            if base_num > 0:
                output.addPage(base_pdf.getPage(base_num))
                if recc_num > 0:
                    print("Found on recc page: " + str(recc_num))
                    output.addPage(recc_pdf.getPage(recc_num))
                break
else:
    for jj in range(base_pdf.getNumPages()):
        output.addPage(cord_pdf.getPage(jj + diff_length))
        output.addPage(base_pdf.getPage(jj))
        if recc_pdf_exists:
            output.addPage(recc_pdf.getPage(jj))

# Finally, output everything to the PDF
# The output name is chosen based on what the name of the coordination file is
output_name = "8.0 - Coordination Results & Recommendations_" + output_name + "2018.pdf"
with open(output_name, "wb") as w:
    output.write(w)
