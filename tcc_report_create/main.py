#!/usr/bin/env python3

from __future__ import print_function

import argparse
import sys

from PyPDF3 import PdfFileWriter, PdfFileReader


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


parser = argparse.ArgumentParser(description='Create TCC Reports correctly.')
parser.add_argument('--cord_path',
                    action='store',
                    help='The path to the Coordination PDF',
                    required=True,
                    type=str)
parser.add_argument('--base_path',
                    action='store',
                    help='The path to the Base TCC PDF',
                    required=True,
                    type=str)
parser.add_argument('--recc_path',
                    action='store',
                    help='The path to the Recommended TCC PDF',
                    required=True,
                    type=str)

opts = vars(parser.parse_args())

output = PdfFileWriter()

cord_in = PdfFileReader(open(opts['cord_path'], 'rb'), False)
base_in = PdfFileReader(open(opts['base_path'], 'rb'), False)
recc_in = PdfFileReader(open(opts['recc_path'], 'rb'), False)

if base_in.getNumPages() != recc_in.getNumPages():
    eprint('Number of pages does not match in the Base and Recommended PDFs')
    exit()

if cord_in.getNumPages() < base_in.getNumPages():
    eprint('Coordination PDF is shorter than the Base/Recc PDFs')
    exit()

diff_length = cord_in.getNumPages() - base_in.getNumPages()
print(diff_length)

for ii in range(diff_length):
    output.addPage(cord_in.getPage(ii))

for jj in range(base_in.getNumPages()):
    output.addPage(cord_in.getPage(jj + diff_length))
    output.addPage(base_in.getPage(jj))
    output.addPage(recc_in.getPage(jj))

with open("TCC_Report.pdf", "wb") as w:
    output.write(w)
