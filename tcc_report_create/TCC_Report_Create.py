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

check_for_rec_list = ['the following settings changes',
                      'shows the effect of recommendations made',
                      'revised TCC']

logging_name = 'TCC_Zipper'
logger = logging.getLogger(logging_name)


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


def find_matching_page(name, pdf_pages, regex, which_pdf):
    for page_num, page in enumerate(pdf_pages, start=0):
        matches = re.finditer(regex, page, re.MULTILINE)
        for match in matches:
            if match.group(2) == name:
                return page_num

    logger.warning('Unable to find a match for %s in %s', name, which_pdf)
    return -1


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
    parser.add_argument('-e', '--empty_path',
                        action='store',
                        help='The path to the empty file that will be inserted if a recommended page cant be found',
                        type=str)
    parser.add_argument('-m', '--matching',
                        action='store_true',
                        help='Whether or not to trigger the TCC name matching method of combining')
    parser.add_argument('-o', '--output',
                        action='store',
                        help='Use to specify the output file path and location',
                        type=str)
    parser.add_argument('-w', '--warning',
                        action='store_true',
                        help='Flag to only log warnings, ignore normal info logs')
    parser.add_argument('-l', '--logging',
                        action='store',
                        help='Specify path of the logging file',
                        type=str)
    # Store the arguments in a dict for easy reference later
    opts = vars(parser.parse_args())
    if bool(opts['default']) and bool(opts['cord_path']):
        eprint('Specified default and manual modes, please choose only one')
        logger.warning('Default and Manual modes selected')
        exit(-9)
    return opts


def glob_checker(glob_list, name):
    root = tk.Tk()
    root.withdraw()

    if len(glob_list) == 0:
        prompt = 'No ' + name + ' PDF found. Please choose preferred'
        path = filedialog.askopenfilename(title=prompt)
        logger.info('No %s PDF found. Prompted for replacement', name)
    elif len(glob_list) > 1:
        prompt = 'Multiple ' + name + ' PDFs found. Please choose preferred.'
        path = filedialog.askopenfilename(title=prompt)
        logger.info('Multiple %s PDF found. Prompted for replacement', name)
    else:
        path = glob_list[0]

    if path == '':
        logger.critical('No %s path chosen')
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
        logger.info(prompt_title)
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
        logger.info('More then one Recommended PDF found')
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
        logger.info('Recommended PDF found')
    else:
        logger.info('Recommended PDF not selected')

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
    empty_path = os.path.join(path, 'PDF', '_INITIAL TCC_Blank_Page.pdf')

    # Build the list of all possible matches to the glob
    base_glob_list = glob.glob(base_path)
    rec_glob_list = glob.glob(rec_path)
    cord_glob_list = glob.glob(cord_path)

    base_path = glob_checker(base_glob_list, 'Base')
    cord_path = glob_checker(cord_glob_list, 'Coordination')
    rec_path, rec_pdf_exists = rec_glob_checker(rec_glob_list)

    if not cord_path or not base_path:
        logging.warning('Unable to find necessary PDF. Exiting')
        exit(-1)

    return cord_path, base_path, rec_path, rec_pdf_exists, empty_path


def manual_mode():
    rec_path = ''
    cord_path = filedialog.askopenfilename(title='Choose preferred Coordination TCC File')
    if cord_path == '':
        prompt = 'Did not provide a path for Coordination'
        logger.critical(prompt)
        eprint(prompt)
        exit(-1)

    base_path = filedialog.askopenfilename(title='Choose preferred Base TCC File')
    if base_path == '':
        prompt = 'Did not provide a path for Base'
        eprint(prompt)
        logger.critical(prompt)
        exit(-1)

    prompt_title = 'Is there a recommended TCC for this project?'
    prompt_message = 'Choose Yes/No'
    result = messagebox.askyesno(title=prompt_title, message=prompt_message)
    if result:
        rec_path = filedialog.askopenfilename(title='Please choose Recommended PDF')
        if rec_path == '':
            rec_pdf_exists = False
            logger.info('No Recommended PDF selected')
        else:
            rec_pdf_exists = True
            logger.info('Recommended PDF selected')
    else:
        rec_pdf_exists = False
        logger.info('No Recommended PDF selected')

    empty_path = filedialog.askopenfilename(title='Locate the TCC_Blank_Page.pdf file')
    if empty_path == '':
        prompt = 'Did not provide a path for TCC File'
        eprint(prompt)
        logger.critical(prompt)
        exit(-1)

    return cord_path, base_path, rec_path, rec_pdf_exists, empty_path


def pdf_selection_via_mode(opts):
    base_path = ''
    rec_path = ''
    cord_path = ''
    rec_pdf_exists = False
    empty_path = ''

    if bool(opts['default']) and not bool(opts['cord_path']):
        logger.info('Default mode selected')
        root = tk.Tk()
        root.withdraw()

        cwd = filedialog.askdirectory(title='Please choose the root of the TCC folders')

        cord_path, base_path, rec_path, rec_pdf_exists, empty_path = default_mode(cwd)

    elif bool(opts['cord_path']) and not bool(opts['default']):
        logger.info('Manual mode selected')
        cord_path = opts['cord_path']
        base_path = opts['base_path']
        rec_path = opts['rec_path']
        rec_pdf_exists = bool(rec_path)
        empty_path = opts['empty_path']

        # Check that if default was NOT selected, we are provided with the proper file flags
        if not (bool(cord_path) and bool(base_path)):
            eprint('Didn\'t specify default, but also didn\'t give all the individual flags',
                   'Either provide all three "--FILE_path" flags OR the "--default" flag')
            logger.critical('Default not selected, but manual flags missing as well')
            exit()

    else:  # Prompt for everything

        root = tk.Tk()
        root.withdraw()
        result = messagebox.askyesno('Use Default Mode?', 'Do you want to use Default Mode to choose files?')
        logger.info('No modes selected, prompting user for choice')
        if result:
            logger.info('Default mode chosen')
            path = filedialog.askdirectory(title='Please choose the root of the TCC Project Folder')
            cord_path, base_path, rec_path, rec_pdf_exists, empty_path = default_mode(path)
        else:
            logger.info('Manual mode chosen')
            cord_path, base_path, rec_path, rec_pdf_exists, empty_path = manual_mode()

    logger.info('\nCord Path: %s\nBase Path: %s\nRec Path: %s\nRecPathExists: %s',
                cord_path, base_path, rec_path, str(rec_pdf_exists))
    return cord_path, base_path, rec_path, rec_pdf_exists, empty_path


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

    logger.info('Chosen output name: %s', output_name)
    return output_name


def zipper(opts, cord_path, base_path, rec_path, rec_pdf_exists, output_name, matching, empty_path):
    # ######### PDF Write Setup ######### #
    # Open the input PDFs
    cord_pdf = PdfFileReader(open(cord_path, 'rb'), False)
    base_pdf = PdfFileReader(open(base_path, 'rb'), False)
    rec_pdf = ''
    if rec_pdf_exists:
        rec_pdf = PdfFileReader(open(rec_path, 'rb'), False)
    empty_pdf = PdfFileReader(open(empty_path, 'rb'), False)

    # Check that the coordination PDF is longer than the base (and therefore rec) pdf too.
    # The Coordination PDF includes pages at the front that do not get sliced in, and instead actually sit
    # in the front. If the Coordination pdf is less than the Base or Rec, these are missing, or there was another error
    if cord_pdf.getNumPages() < base_pdf.getNumPages():
        prompt = 'Coordination PDF is shorter than the Base PDF'
        eprint(prompt)
        logger.critical(prompt)
        exit(-7)

    # Find the difference in length of the PDFs, these are the leader pages of the coordination
    diff_length = cord_pdf.getNumPages() - base_pdf.getNumPages()
    logger.info('Diff Length: %s', str(diff_length))

    output = PdfFileWriter()

    for ii in range(diff_length):
        output.addPage(cord_pdf.getPage(ii))

    if matching:

        logger.info("Converting Coordination PDF to string")
        logging.disable(logging.INFO)
        cord_str_pages = pdf_pages_to_list_of_strings(cord_path)
        logging.disable(logging.NOTSET)

        logger.info("Converting Base PDF to string")
        logging.disable(logging.INFO)
        base_str_pages = pdf_pages_to_list_of_strings(base_path)
        logging.disable(logging.NOTSET)

        rec_str_pages = []
        if rec_pdf_exists:
            logging.disable(logging.INFO)
            logger.info("Converting Recommended PDF to string")
            rec_str_pages = pdf_pages_to_list_of_strings(rec_path)
            logging.disable(logging.NOTSET)

        regex_cord = r'(TCC Curve: )(TCC_[\d]+[a-zA-Z]?)([-_#$\w\d\[\] ]*)'
        regex_base_rec = r'(TCC Name: )(TCC_[\d]+[a-zA-Z]?)([-_#$\w\d\[\] ]*)'

        for ii in range(diff_length, len(cord_str_pages)):
            output.addPage(cord_pdf.getPage(ii))
            tcc_matches = re.finditer(regex_cord, cord_str_pages[ii], re.MULTILINE)

            for match_num, tcc_match in enumerate(tcc_matches, start=1):
                tcc_name = tcc_match.group(2)
                logger.info("Attempting to find: " + tcc_name)
                base_num = find_matching_page(tcc_name, base_str_pages, regex_base_rec, 'Base PDF')
                if base_num != -1:
                    logger.info('Found on base page: %s', str(base_num))
                rec_page_flag = check_for_rec(cord_str_pages[ii])
                rec_num = 0
                if rec_pdf_exists and rec_page_flag:
                    rec_num = find_matching_page(tcc_name, rec_str_pages, regex_base_rec, 'Rec PDF')
                    if rec_num != -1:
                        logger.info('Found on rec page: %s', str(rec_num))
                    else:
                        output.addPage(empty_pdf.getPage(0))
                if base_num > 0:
                    output.addPage(base_pdf.getPage(base_num))
                    if rec_num > 0:
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
    if opts['output']:
        output_name = opts['output']
    else:
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
                                      message='Matches by TCC title. Takes a little longer')

    if ret_val:
        logger.info('Matching option chosen')
    else:
        logger.info('Zipper option chosen')

    return ret_val


def valid_tcc_name(tcc_name):
    valid_tcc_name_pattern = re.compile(r'TCC_[\d]+_[\w/ \[\]\"-]+')
    if not valid_tcc_name_pattern.match(tcc_name):
        logger.critical('Invalid TCC Name: %s', tcc_name)


def check_for_rec(cord_str_page):
    for tag in check_for_rec_list:
        if tag in cord_str_page:
            logger.info('Found Rec phrase in Coordination')
            return True
    logger.info('Unable to find Rec phrase in Coordination')
    return False


def main():
    # ######### Argument Handling ######### #
    opts = argument_handler()

    if opts['warning']:
        logging_level = logging.WARNING
    else:
        logging_level = logging.INFO

    if opts['logging']:
        log_base_path = opts['logging']
    else:
        log_base_path = os.path.dirname(os.path.realpath(__file__))
        if not os.path.isdir('Logs'):
            os.mkdir(os.path.join(log_base_path, 'Logs'))
        log_base_path = os.path.join(log_base_path, 'Logs', 'TCC_Create_Logs.log')

    logging.basicConfig(filename=log_base_path,
                        level=logging_level,
                        format='%(levelname)s|%(name)s|%(message)s',
                        filemode='w')

    # ######### Path Handling ######### #
    cord_path, base_path, rec_path, rec_pdf_exists, empty_path = pdf_selection_via_mode(opts)

    # ######### PDF Output Name Selection ######### #
    output_name = output_name_selector(cord_path)

    matching = do_matching_check(opts)

    # ######### PDF Write Setup ######### #
    zipper(opts, cord_path, base_path, rec_path, rec_pdf_exists, output_name, matching, empty_path)


if __name__ == '__main__':
    main()
