from PyPDF2 import PdfReader
import win32com.client
import datetime
import glob
import ocrmypdf
import os
import sqlite3

#options 

company_name = "company"

#use if you are having with random spaces between words
do_ocr = True 

#end options





#debug options

#delete ocr_pdf after use
delete_ocr_pdf = True

#delete db after use
delete_db = True

#end debug options

if os.path.exists("files.db"):
    os.remove("files.db")
if os.path.exists("files.db-journal"):
    os.remove("files.db-journal")

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
    if os.path.exists(f'temp/ocr_{file}'):
        return f'temp/ocr_{file}'
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

    for attach in attachments:
        newmail.Attachments.Add(file_dir+"\\"+attach)

    newmail.Display() 

def GetDetails(lines):
    return ExtractName(lines),ExtractEmailAddress(lines)

def GetLines(pdf):
    return SplitPdf(GetText(converted_pdf))


def SendEmails(list_of_unique_emails):
    for tup_email in list_of_unique_emails :
        email = tup_email[0]
        name = tup_email[1]

        attachments = []

        sql_attachments = cur.execute("SELECT pdf FROM FILES WHERE email = ? AND name = ?", (email,name,),).fetchall()
        for attachment in sql_attachments:
            attachments.append(attachment[0])

        SendEmail(name,email,attachments)




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


    list_of_unique_emails = cur.execute("SELECT DISTINCT email,name FROM FILES").fetchall()

    SendEmails(list_of_unique_emails)

    if delete_ocr_pdf:
        DeleteTempPdf()
    if errors != "":
        print("ERRORS:")
        print(errors)
    con.commit()
    cur.close()
    con.close()
    if delete_db:
        import time
        print("waiting for delete")
        time.sleep(5)
        if os.path.exists("files.db"):
            os.remove("files.db")
        if os.path.exists("files.db-journal"):
            os.remove("files.db-journal")
    






    


