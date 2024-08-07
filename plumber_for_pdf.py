import pdfplumber
import glob

def GetPdf():
    pdfs = []
    for file in glob.glob("*.pdf"):
        pdfs.append(file)
    return pdfs


def process_pdfs(pdfs):
    print("------------------")
    for pdf in pdfs:
        emails = []
        names = []
        with pdfplumber.open(pdf) as pdf:
            first_page = pdf.pages[0]
            words = first_page.extract_words()
            for i in range(len(first_page.extract_words())):
                if words[i]['text'] == "Email" and words[i+1]['text'] == "address:":
                    for x in range(i+2, len(words)):
                        if words[x]['text'] == "VAT" or words[x]['text'] == "Nr:" or words[x]['text'] == "FINANCIAL" or words[x]['text'] == "PERIOD" or words[x]['text'] == "15%":
                            break
                        potential_email = words[x]['text'].strip().rstrip(";")
                        if "@" in potential_email:
                            emails.append(potential_email)
                if words[i]['text'] == "To:":
                    name = ""
                    for x in range(i+1, len(words)):
                        if words[x]['text'].strip() == "From:":
                            names.append(name.strip())
                            break
                        name += words[x]['text'].strip() + " "
                    
        print(emails)
        print(names)
        print("------------------")
        

    
if __name__ == "__main__":
    pdfs = GetPdf()
    process_pdfs(pdfs)