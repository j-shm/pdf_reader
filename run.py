from PyPDF2 import PdfReader
import win32com.client
import datetime
import glob
import ocrmypdf
import os
import sqlite3

company_name = "company"
do_ocr = True #use if you are having with random spaces between words

file_dir = os.getcwd()


con = sqlite3.connect("files.db")
cur = con.cursor()#
table = """
CREATE TABLE IF NOT EXISTS FILES (
	pdf text PRIMARY KEY,
    email text,
	name text
);
"""
cur.execute(table)

def ConvertPdf(file) -> str:
    """open the file(pdf) and return the text"""
    if not do_ocr:
        return file
    try:
        ocrmypdf.ocr(file, f'temp/ocr_{file}',force_ocr = True)
    except Exception as e:
        print(e)
        return None
    return f'temp/ocr_{file}'

def GetText(file, page = 0):
    """Get the text from the pdf"""
    if file == None:
        return None
    reader = PdfReader(file)
    page = reader.pages[page]
    return page.extract_text()

def DeleteTempPdf():
    """Delete the temporary ocr pdf"""
    files = glob.glob('temp/*')
    for f in files:
        os.remove(f)

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

def ExtractEmailAddress(lines):
    """Get email from pdf"""
    for line in lines:
        splitlines = line.split(":")
        for index, splitline in enumerate(splitlines):
            if splitline.strip() == "Email Address":
                return splitlines[index+1].strip()
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
    newmail.Subject = f'{company_name} - {GetDate()} - {name.split(".")[0]}'
    newmail.To=''+email
    newmail.CC=''+email

    for attach in attachments:
        newmail.Attachments.Add(file_dir+"\\"+attach)

    newmail.Display() 

def GetDetails(lines):
    return ExtractName(lines),ExtractEmailAddress(lines)

def GetLines(pdf):
    return SplitPdf(GetText(converted_pdf))

if __name__ == '__main__':
    errors = ""
    pdfs = GetPdf()
    for pdf in pdfs:

        converted_pdf = ConvertPdf(pdf)
        if converted_pdf == None:
            errors += f'{pdf} failed'
            continue
        
        lines = GetLines(converted_pdf)
        name,email = GetDetails(lines)
        
        cur.execute(f'INSERT INTO FILES(pdf,email,name) VALUES(?,?,?)',(pdf,email,name))


        #excel = GetExcel(pdf)

        #attachments = [pdf]
        #if excel != None:
        #    attachments.append(excel)

        #SendEmail(name,email,attachments)

        


    DeleteTempPdf()
    if errors != "":
        print("ERRORS:")
        print(errors)
    con.commit()
    cur.close()
    con.close()





def functions():
    """just stuff that will be needed"""

    #for getting all the details from an email
    temp_email = "bob@gmail.com"
    for result in cur.execute("SELECT * FROM FILES WHERE email = ?", (temp_email,),).fetchall():
        print(result)
        print(" ")

    #finding out all unique emails
    print("new statment :)")
    for result in cur.execute("SELECT DISTINCT email FROM FILES").fetchall():
        print(result[0])
        print(" ")
    


