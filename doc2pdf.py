# coding: utf-8

import os
import shutil
import win32com
from win32com.client import *


def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)


def doc2pdf(doc_fullpath, word):
    doc_fullpath = doc_fullpath.replace("/", "\\")
    print(doc_fullpath)

    dirname = os.path.dirname(doc_fullpath)
    current_file = os.path.basename(doc_fullpath)
    fname, ext = os.path.splitext(current_file)
    doc = word.Documents.Open(doc_fullpath)
    # Save as PDF file
    doc.SaveAs(dirname + '/' + fname + '.pdf', FileFormat=17)
    doc.Close()


if __name__ == '__main__':
    s = input("Dir: ")
    root_dir = s.strip('\"')
    root_dir_copy = root_dir + '__copy'
    shutil.copytree(root_dir, root_dir_copy)

    # com object
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    print('Processing...')

    for i in all_files(root_dir_copy):
        dirname = os.path.dirname(i)
        current_file = os.path.basename(i)
        fname, ext = os.path.splitext(current_file)

        if ext == '.doc' or ext == '.docx':
            try:
                # Convert Doc(x) to PDF
                doc2pdf(dirname + '/' + current_file, word)
            except:
                print('Error: ' + i)

            # Delete doc(x) files
            os.remove(dirname + '/' + current_file)
        else:
            # Delete non-PDF files
            os.remove(i)

    word.Quit()

    print('')
    print('Done!')
    print('Enter to exit.')
    os.system("pause > nul")
