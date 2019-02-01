#!/usr/bin/env python3
# ######### IMPORTS ######### #
from __future__ import print_function

import argparse
import glob
import os
import re
import sys
import tkinter as tk
from io import StringIO
from tkinter import filedialog, messagebox

from PyPDF3 import PdfFileWriter, PdfFileReader
from pdfminer3.converter import TextConverter
from pdfminer3.layout import LAParams
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.pdfpage import PDFPage


# ######### ERROR HElPER ######### #
def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


def pdf_pages_to_list_of_strings(pdf_path):
    pdf = open(pdf_path, 'rb')
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # Process each page contained in the document.

    pages_text = []

    pdf_pages = PDFPage.get_pages(pdf)

    for page in pdf_pages:
        # Get (and store) the "cursor" position of stream before reading from PDF
        # On the first page, this will be zero
        read_position = retstr.tell()

        # Read PDF page, write text into stream
        interpreter.process_page(page)

        # Move the "cursor" to the position stored
        retstr.seek(read_position, 0)

        # Read the text (from the "cursor" to the end)
        page_text = retstr.read()

        # Add this page's text to a convenient list
        pages_text.append(page_text)

    return pages_text


def find_matching_page(name, pdf_pages, regex):
    for page_num, page in enumerate(pdf_pages, start=0):
        matches = re.finditer(regex, page, re.MULTILINE)
        for match in matches:
            if match.group(2) == name:
                return page_num
    return 0


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
parser.add_argument('-r', '--recc_path',
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
recc_path = ''

# Empty lists to reference later
cord_glob_list = []
recc_glob_list = []
base_glob_list = []
recc_pdf_exists = False

# If the default flag was not selected
if not opts['default']:
    # Fill the paths. If they weren't provided they will be null
    cord_path = opts['cord_path']
    base_path = opts['base_path']
    recc_path = opts['recc_path']

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
    recc_path = os.path.join(cwd, 'TCCs', rec_glob)

    # Build the list of all possible matches to the glob
    base_glob_list = glob.glob(base_path)
    recc_glob_list = glob.glob(recc_path)
    cord_glob_list = glob.glob(cord_path)

    recc_pdf_exists = len(recc_glob_list) > 0

# If there is more than one, ask them to remove the extra
if len(base_glob_list) > 1:

    root = tk.Tk()
    root.withdraw()

    base_path = filedialog.askopenfilename(title='Choose preferred Base TCC File')
    if base_path == '':
        eprint('Did not provide a path for Base')
        exit(-1)
# If there aren't any files in the folder, throw error
elif len(base_glob_list) == 0:
    root = tk.Tk()
    root.withdraw()

    base_path = filedialog.askopenfilename(title='Please find and select preferred Base TCC File')
    if base_path == '':
        eprint('Did not provide a file for Base')
        exit(-2)
else:
    base_path = base_glob_list[0]

if not recc_pdf_exists:
    recc_pdf_exists = messagebox.askyesno('Recommended TCC?',
                                          'Is there a missing Recommended TCC you would like to use?')

if recc_pdf_exists:
    if len(recc_glob_list) > 1:
        root = tk.Tk()
        root.withdraw()

        recc_path = filedialog.askopenfilename(title='Choose preferred Recc TCC File')
        if recc_path == '':
            eprint('Did not provide a path for Base')
            exit(-3)
    elif len(recc_glob_list) == 0:
        root = tk.Tk()
        root.withdraw()

        recc_path = filedialog.askopenfilename(title='Choose preferred Recc TCC File')
        if recc_path == '':
            eprint('Did not provide a path for Base')
            exit(-4)
    else:
        recc_path = recc_glob_list[0]

# Finally, the same checks happen for the coordination file

if len(cord_glob_list) > 1:
    root = tk.Tk()
    root.withdraw()

    cord_path = filedialog.askopenfilename(title='Choose preferred Coordination File')
    if cord_path == '':
        eprint('Did not provide a path for Coordination')
        exit(-5)
elif len(cord_glob_list) == 0:
    root = tk.Tk()
    root.withdraw()

    cord_path = filedialog.askopenfilename(title='Choose preferred Coordination File')
    if cord_path == '':
        eprint('Did not provide a path for Coordination')
        exit(-6)
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
    recc_pdf = PdfFileReader(open(recc_path, 'rb'), False)

# Check that the same number of pages exist in the Base and Recommended PDFs.
# They should be the same length
if recc_pdf_exists:
    if base_pdf.getNumPages() != recc_pdf.getNumPages():
        eprint('Number of pages does not tcc_matches in the Base and Recommended PDFs')
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
        recc_str_pages = pdf_pages_to_list_of_strings(recc_path)

    # regex_cord = re.compile(r"\bTCC_\*")
    regex_cord = r"(TCC Curve: )(TCC_[\w/ \[\]\"-]+)"
    regex_base_recc = r"(TCC Name: )(TCC_[\w/ \[\]\"-]+)"

    for ii in range(diff_length, len(cord_str_pages)):
        output.addPage(cord_pdf.getPage(ii))
        tcc_matches = re.finditer(regex_cord, cord_str_pages[ii], re.MULTILINE)
        for match_num, tcc_match in enumerate(tcc_matches, start=1):
            recc_num = 0
            if recc_pdf_exists:
                recc_num = find_matching_page(tcc_match.group(2), recc_str_pages, regex_base_recc)
            print("Attempting to find: " + tcc_match.group(2))
            base_num = find_matching_page(tcc_match.group(2), base_str_pages, regex_base_recc)
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
output_name = "8.0 - Coordination Results & Recommendations_" + output_name + "2018_NEW.pdf"
output_name = os.path.join(os.getcwd(), 'PDF', output_name)
with open(output_name, "wb") as w:
    output.write(w)
