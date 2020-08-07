import os
from os.path import join

from win32com.client import Dispatch
from os import walk

wdFormatPDF = 17


def doc2pdf(input_file):
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(input_file.replace(".doc", ".pdf"), FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


if __name__ == "__main__":
    doc_files = []
    directory = "C:\\Users\\localhost\\Desktop"
    doc2pdf(r"C:\Users\localhost\Desktop\aa.doc")
    for root, dirs, filenames in walk(directory):
        for file in filenames:
            # print(root, dirs, file)
            print(join(root, file))
    print("_" * 120)
    for file in os.listdir(r'C:\Users\localhost\Desktop'):
        if os.path.isfile(join(r'C:\Users\localhost\Desktop', file)):
            print(file)