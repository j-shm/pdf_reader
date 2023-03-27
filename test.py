# use this to see the data for each pdf to see if you need to enable ocr
# used for ease of setting up with people who aren't familiar with python
# (thrown together with haste dont refactor only for testing)

import glob
from PyPDF2 import PdfReader

if __name__ == '__main__':
    pdfs = []
    for file in glob.glob("*.pdf"):
        pdfs.append(file)
    for count,pdf in enumerate(pdfs):
        print("-=" + pdf + "=-");

        reader = PdfReader(pdf)
        page = reader.pages[0]
        text = page.extract_text().split("\n")
        print(repr(text))

        print("-==-")

        if count == 5:
            print("look for random spaces")
            print("if there are random spaces...")
            print("...where there should not be use ocr")
            print("example:")
            print("To: Ben Jacobs From:")
            print("Email address:  ben@gmail.com VAT Nr:  ")
            exit()
    print("look for random spaces")
    print("if there are random spaces...")
    print("...where there should not be use ocr")
    print("example:")
    print("To: Ben Jacobs From:")
    print("Email address:  ben@gmail.com VAT Nr:  ")
    print("...where there should not be use ocr")