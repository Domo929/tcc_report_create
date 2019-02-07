#!/usr/bin/env python3
# ######### IMPORTS ######### #
from __future__ import print_function

import argparse
import glob
import logging
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

    logging.warning('Unable to find a match for %s', name)
    return 0


# ######### ARGUMENT HANDLER ######### #
def argument_handler():
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
    if bool(opts['default']) and bool(opts['cord_path']):
        eprint('Specified default and manual modes, please choose only one')
        logging.warning('Default and Manual modes selected')
        exit(-9)
    return opts


def glob_checker(glob_list, name):
    root = tk.Tk()
    root.withdraw()

    if len(glob_list) == 0:
        prompt = 'No ' + name + ' PDF found. Please choose preferred'
        path = filedialog.askopenfilename(title=prompt)
        logging.info('No %s PDF found. Prompted for replacement', name)
    elif len(glob_list) > 1:
        prompt = 'Multiple ' + name + ' PDFs found. Please choose preferred.'
        path = filedialog.askopenfilename(title=prompt)
        logging.info('Multiple %s PDF found. Prompted for replacement', name)
    else:
        path = glob_list[0]

    if path == '':
        logging.critical('No %s path chosen')
    return path


def rec_glob_checker(glob_list):
    root = tk.Tk()
    root.withdraw()
    pdf_exists = False
    path = ''

    if len(glob_list) == 0:
        prompt_title = 'No Recommended PDF found'
        prompt_message = 'Does one exist for this project?'
        result = messagebox.askyesno(title=prompt_title, message=prompt_message)
        logging.info(prompt_title)
        if result == 'yes':
            path = filedialog.askopenfilename(title='Please choose Recommended PDF')
            if path == '':
                pdf_exists = False
            else:
                pdf_exists = True
        else:
            pdf_exists = False
    elif len(glob_list) > 1:
        prompt_title = 'More than one Recommended PDF found, please choose preferred'
        path = filedialog.askopenfilename(title=prompt_title)
        logging.info('More then one Recommended PDF found')
        if path == '':
            pdf_exists = False
        else:
            pdf_exists = True

    else:
        path = glob_list[0]
        if path == '':
            pdf_exists = False
        else:
            pdf_exists = True

    if pdf_exists:
        logging.info('Recommended PDF found')
    else:
        logging.info('Recommended PDF not selected')

    return path, pdf_exists


def default_mode(path):
    # This allows us to find files that have different versions
    base_glob = 'TCC?Base?v*.pdf'
    rec_glob = 'TCC?Rec?v*.pdf'
    cord_glob = '*Coordination*.pdf'

    # Build the full path in order to start searching
    cord_path = os.path.join(path, 'PDF', cord_glob)
    base_path = os.path.join(path, 'TCCs', base_glob)
    rec_path = os.path.join(path, 'TCCs', rec_glob)

    # Build the list of all possible matches to the glob
    base_glob_list = glob.glob(base_path)
    rec_glob_list = glob.glob(rec_path)
    cord_glob_list = glob.glob(cord_path)

    base_path = glob_checker(base_glob_list, 'Base')
    cord_path = glob_checker(cord_glob_list, 'Coordination')
    rec_path, rec_pdf_exists = rec_glob_checker(rec_glob_list)

    return cord_path, base_path, rec_path, rec_pdf_exists


def manual_mode():
    rec_path = ''
    cord_path = filedialog.askopenfilename(title='Choose preferred Coordination TCC File')
    if cord_path == '':
        prompt = 'Did not provide a path for Coordination'
        logging.critical(prompt)
        eprint(prompt)
        exit(-1)

    base_path = filedialog.askopenfilename(title='Choose preferred Base TCC File')
    if base_path == '':
        prompt = 'Did not provide a path for Base'
        eprint(prompt)
        logging.critical(prompt)
        exit(-1)

    prompt_title = 'Is there a recommended TCC for this project?'
    prompt_message = 'Choose Yes/No'
    result = messagebox.askyesno(title=prompt_title, message=prompt_message)
    if result:
        rec_path = filedialog.askopenfilename(title='Please choose Recommended PDF')
        if rec_path == '':
            rec_pdf_exists = False
            logging.info('No Recommended PDF selected')
        else:
            rec_pdf_exists = True
            logging.info('Recommended PDF selected')
    else:
        rec_pdf_exists = False
        logging.info('No Recommended PDF selected')

    return cord_path, base_path, rec_path, rec_pdf_exists


def pdf_selection(opts):
    base_path = ''
    rec_path = ''
    cord_path = ''
    rec_pdf_exists = False

    if bool(opts['default']) and not bool(opts['cord_path']):
        logging.info('Default mode selected')
        root = tk.Tk()
        root.withdraw()

        cwd = filedialog.askdirectory(title='Please choose the root of the TCC folders')

        cord_path, base_path, rec_path, rec_pdf_exists = default_mode(cwd)

    elif bool(opts['cord_path']) and not bool(opts['default']):
        logging.info('Manual mode selected')
        cord_path = opts['cord_path']
        base_path = opts['base_path']
        rec_path = opts['rec_path']
        rec_pdf_exists = bool(rec_path)

        # Check that if default was NOT selected, we are provided with the proper file flags
        if not (bool(cord_path) and bool(base_path)):
            eprint('Didn\'t specify default, but also didn\'t give all the individual flags',
                   'Either provide all three "--FILE_path" flags OR the "--default" flag')
            logging.critical('Default not selected, but manual flags missing as well')
            exit()

    else:  # Prompt for everything

        root = tk.Tk()
        root.withdraw()
        result = messagebox.askyesno('Use Default Mode?', 'Do you want to use Default Mode to choose files?')
        logging.info('No modes selected, prompting user for choice')
        if result:
            logging.info('Default mode chosen')
            path = filedialog.askdirectory(title='Please choose the root of the TCC Project Folder')
            cord_path, base_path, rec_path, rec_pdf_exists = default_mode(path)
        else:
            logging.info('Manual mode chosen')
            cord_path, base_path, rec_path, rec_pdf_exists = manual_mode()

    logging.info('Cord Path: %s\nBase Path: %s\nRec Path: %s\nRecPathExists: %s',
                 cord_path, base_path, rec_path, str(rec_pdf_exists))
    return cord_path, base_path, rec_path, rec_pdf_exists


def output_name_selector(cord_path):
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

    logging.info('Chosen output name: %s', output_name)
    return output_name


def zipper(cord_path, base_path, rec_path, rec_pdf_exists, output_name, matching):
    # ######### PDF Write Setup ######### #
    # Open the input PDFs
    cord_pdf = PdfFileReader(open(cord_path, 'rb'), False)
    base_pdf = PdfFileReader(open(base_path, 'rb'), False)
    rec_pdf = ''
    if rec_pdf_exists:
        rec_pdf = PdfFileReader(open(rec_path, 'rb'), False)

    # Check that the coordination PDF is longer than the base (and therefore rec) pdf too.
    # The Coordination PDF includes pages at the front that do not get sliced in, and instead actually sit
    # in the front. If the Coordination pdf is less than the Base or Rec, these are missing, or there was another error
    if cord_pdf.getNumPages() < base_pdf.getNumPages():
        prompt = 'Coordination PDF is shorter than the Base PDF'
        eprint(prompt)
        logging.critical(prompt)
        exit(-7)

    # Find the difference in length of the PDFs, these are the leader pages of the coordination
    diff_length = cord_pdf.getNumPages() - base_pdf.getNumPages()
    logging.info('Diff Length: %s', str(diff_length))

    output = PdfFileWriter()

    for ii in range(diff_length):
        output.addPage(cord_pdf.getPage(ii))

    if matching:
        logging.info("Converting Coordination PDF to string")
        cord_str_pages = pdf_pages_to_list_of_strings(cord_path)

        logging.info("Converting Base PDF to string")
        base_str_pages = pdf_pages_to_list_of_strings(base_path)

        rec_str_pages = []
        if rec_pdf_exists:
            logging.info("Converting Recommended PDF to string")
            rec_str_pages = pdf_pages_to_list_of_strings(rec_path)

        # regex_cord = re.compile(r"\bTCC_\*")
        regex_cord = r"(TCC Curve: )(TCC_[\w/ \[\]\"-]+)"
        regex_base_rec = r"(TCC Name: )(TCC_[\w/ \[\]\"-]+)"

        for ii in range(diff_length, len(cord_str_pages)):
            output.addPage(cord_pdf.getPage(ii))
            tcc_matches = re.finditer(regex_cord, cord_str_pages[ii], re.MULTILINE)

            for match_num, tcc_match in enumerate(tcc_matches, start=1):
                rec_num = 0
                if rec_pdf_exists:
                    rec_num = find_matching_page(tcc_match.group(2), rec_str_pages, regex_base_rec)
                logging.info("Attempting to find: " + tcc_match.group(2))
                base_num = find_matching_page(tcc_match.group(2), base_str_pages, regex_base_rec)
                logging.info('Found on base page: %s', str(base_num))
                if base_num > 0:
                    output.addPage(base_pdf.getPage(base_num))
                    if rec_num > 0:
                        logging.info('Found on rec page: %s', str(rec_num))
                        output.addPage(rec_pdf.getPage(rec_num))
                    break
    else:
        for jj in range(base_pdf.getNumPages()):
            output.addPage(cord_pdf.getPage(jj + diff_length))
            output.addPage(base_pdf.getPage(jj))
            if rec_pdf_exists:
                output.addPage(rec_pdf.getPage(jj))

    # Finally, output everything to the PDF
    # The output name is chosen based on what the name of the coordination file is
    output_name = "8.0 - Coordination Results & Recommendations_" + output_name + "2018_NEW.pdf"
    output_name = os.path.join(os.path.dirname(os.path.abspath(cord_path)), output_name)
    with open(output_name, "wb") as w:
        output.write(w)


def do_matching_check(opts):
    ret_val = False
    if opts['matching']:
        ret_val = True
    elif bool(opts['default']) or bool(opts['cord_path']):
        ret_val = False
    else:
        root = tk.Tk()
        root.withdraw()

        ret_val = messagebox.askyesno(title='Do you want to perform matching?',
                                      message='Matches by TCC title. Takes a long time')

    if ret_val:
        logging.info('Matching option chosen')
    else:
        logging.info('Zipper option chosen')

    return ret_val


def main():
    logging.basicConfig(filename='TCC_Create_Logs.log',
                        level=logging.INFO,
                        format='%(levelname)s:%(message)s',
                        filemode='w')

    # ######### Argument Handling ######### #
    opts = argument_handler()

    # ######### Path Handling ######### #
    cord_path, base_path, rec_path, rec_pdf_exists = pdf_selection(opts)

    # ######### PDF Output Name Selection ######### #
    output_name = output_name_selector(cord_path)

    matching = do_matching_check(opts)

    # ######### PDF Write Setup ######### #
    zipper(cord_path, base_path, rec_path, rec_pdf_exists, output_name, matching)


if __name__ == '__main__':
    main()
