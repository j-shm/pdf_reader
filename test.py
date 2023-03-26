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
        print("look for random spaces")
        print("if there are random spaces...")
        print("...where there should not be use ocr")

        print("-=" + pdf + "=-");

        reader = PdfReader(file)
        page = reader.pages[page]
        text = page.extract_text().split("\n")
        print(text)

        print("-=" + pdf + "=-")

        if count == 5:
            print("look for random spaces")
            print("if there are random spaces...")
            print("...where there should not be use ocr")
            exit()