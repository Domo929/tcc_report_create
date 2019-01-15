#!/usr/bin/env python3

from __future__ import print_function

import argparse
import glob
import os
import sys

from PyPDF3 import PdfFileWriter, PdfFileReader


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


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
parser.add_argument('--recc_path',
                    action='store',
                    help='The path to the Recommended TCC PDF',
                    type=str)

opts = vars(parser.parse_args())

cord_path = ''
base_path = ''
recc_path = ''

if not opts['default']:
    cord_path = opts['cord_path']
    base_path = opts['base_path']
    recc_path = opts['recc_path']

    if not (cord_path and base_path and recc_path):
        eprint('Didn\'t specify default, but also didn\'t give all the individual flags')
        exit()

else:
    cwd = os.getcwd()

    base_glob = 'TCC-Base_v*.pdf'
    recc_glob = 'TCC-Rec_v*.pdf'

    cord_path = os.path.join(cwd, 'PDF', 'Coordination.pdf')
    base_path = os.path.join(cwd, 'TCCs', base_glob)
    recc_path = os.path.join(cwd, 'TCCs', recc_glob)

    base_glob_list = glob.glob(base_path)
    if len(base_glob_list) > 1:
        eprint('Found multiple versions of TCC_Base. Please remove all extras')
        exit()
    elif len(base_glob_list) == 0:
        eprint('Found no version of TCC_base. Please provide a sheet')
        exit()
    else:
        base_path = base_glob_list[0]

    recc_glob_list = glob.glob(recc_path)
    if len(recc_glob_list) > 1:
        eprint('Found multiple versions of TCC_Rec. Please remove all extras')
        exit()
    elif len(recc_glob_list) == 0:
        eprint('Found no version of TCC_rec. Please provide a sheet')
        exit()
    else:
        recc_path = recc_glob_list[0]


output = PdfFileWriter()

cord_in = PdfFileReader(open(cord_path, 'rb'), False)
base_in = PdfFileReader(open(base_path, 'rb'), False)
recc_in = PdfFileReader(open(recc_path, 'rb'), False)

if base_in.getNumPages() != recc_in.getNumPages():
    eprint('Number of pages does not match in the Base and Recommended PDFs')
    exit()

if cord_in.getNumPages() < base_in.getNumPages():
    eprint('Coordination PDF is shorter than the Base/Recc PDFs')
    exit()

diff_length = cord_in.getNumPages() - base_in.getNumPages()

for ii in range(diff_length):
    output.addPage(cord_in.getPage(ii))

for jj in range(base_in.getNumPages()):
    output.addPage(cord_in.getPage(jj + diff_length))
    output.addPage(base_in.getPage(jj))
    output.addPage(recc_in.getPage(jj))

with open("8.0 - Coordination Results & Recommendations_CE2018.pdf", "wb") as w:
    output.write(w)
