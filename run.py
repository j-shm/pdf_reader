from PyPDF2 import PdfReader
import win32com.client
import datetime
import glob
import ocrmypdf

file_dir = "E:\\projects\\pdf_reader"



# we should really be using ocr to make sure it acc works
def ConvertPdf(file) -> str:
    """open the file(pdf) and return the text"""
    try:
        ocrmypdf.ocr(file, f'ocr_{file}',force_ocr = True)
    except Exception as e:
        print(e)
        return None
    return f'ocr_{file}'

def GetText(file):
    if file == None:
        return None
    reader = PdfReader(file)
    page = reader.pages[0]
    return page.extract_text()

def DeletePdf(file):
    

def SplitPdf(text):
    """Split the pdf into lines"""
    return text.split("\n")

def ExtractName(lines):
    """Get name from pdf"""
    for line in lines:
        splitlines = line.split(":")
        for index, splitline in enumerate(splitlines):
            if splitline == "To":
                return splitlines[index+1].strip()
    return ""

#sometimes multiple seperated by ;
def ExtractEmailAddress(lines):
    """Get email from pdf"""
    for line in lines:
        splitlines = line.split(":")
        for index, splitline in enumerate(splitlines):
            if splitline.strip() == "Email address":
                print(splitlines[index])
    return ""

def GetDate():
    """Get month and year from pdf"""
    now = datetime.datetime.now()
    month = now.strftime('%B')
    year = now.year
    return f'{month} {year}'

def GetPdf():
    """Get all pdf in dir"""
    pdfs = []
    for file in glob.glob("*.pdf"):
        pdfs.append(file)
    return pdfs

def GetExcel(name):
    """Get excel for the pdf"""
    tname = name.split(".")[0]
    file = glob.glob(f'{tname}.xsls')
    if file:
        return file[0]
    
    xlsxfile = glob.glob(f'{tname}.xlsx')
    if xlsxfile:
        return xlsxfile[0]
    
    return None




def SendEmail(name, email, attachments):
    """generate email"""
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject = 'Company - '+GetDate()+" - "+name.split(".")[0]
    newmail.To=''+email
    newmail.CC=''+email

    for attach in attachments:
        newmail.Attachments.Add(file_dir+attach)

    newmail.Display() 

if __name__ == '__main__':
    pdfs = GetPdf()
    for pdf in pdfs:
        text = GetText(ConvertPdf(pdf))
        lines = SplitPdf(text)
        for line in lines:
            print(lines)
        #name = ExtractName(lines)
        #email = ExtractEmailAddress(lines)
        #excel = GetExcel(pdf)
        #attachments = [pdf]
        #if excel != None:
        #    attachments.append(excel)
        #SendEmail(name,email,attachments)

    


