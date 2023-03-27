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
do_ocr = False 

one_at_a_time = True
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
                return splitlines[index+1].strip()[:-5]
    return ""

def ExtractEmailAddress(lines):
    """Get email from pdf"""
    for line in lines:
        if "Email address:" in line:
            splitline = line.split(":")[1].strip()[:-7]
            return splitline
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
    excel_files_1 = glob.glob("*.xsls")
    excel_files_2 = glob.glob("*.xlsx")
    excel_files = excel_files_1 + excel_files_2
    matches = [string for string in excel_files if tname in string]
    return matches




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

def GetLines(converted_pdf):
    return SplitPdf(GetText(converted_pdf))


def SendEmails(list_of_unique_emails):
    #used if there is multiple pdfs to one email
    for tup_email in list_of_unique_emails :
        email = tup_email[0]
        name = tup_email[1]

        attachments = []
        
        sql_attachments = cur.execute("SELECT pdf FROM FILES WHERE email = ? AND name = ?", (email,name,),).fetchall()
        for attachment in sql_attachments:
            attachments.append(attachment[0])
            excels = GetExcel(attachment[0])
            for excel in excels:
                if excel not in attachments:
                    attachments.append(excel) 
        SendEmail(name,email,attachments)

def SendEmails():   
    items = cur.execute("SELECT * FROM FILES").fetchall()
    for item in items:
        if one_at_a_time:
            print("")
            print("Enter anything to go to email: \nDetails:\n" + "From: " + company_name + "\n" + item[0] + "\n" + item[1] + "\n" + item[2])
            input()
        attachments = [item[0]]
        email = item[1]
        name = item[2]
        excels = GetExcel(item[0].split(".")[0])
        for excel in excels:
            if excel not in attachments:
                attachments.append(excel) 
        SendEmail(name,email,attachments)

if __name__ == '__main__':
    company_name = input("Enter company name: ")
    print("COMPANY NAME: " + company_name)
    print("")
    selection = input("Type y to continue:")
    if selection != "y":
        exit()

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


    #list_of_unique_emails = cur.execute("SELECT DISTINCT email,name FROM FILES").fetchall()

    #SendEmails(list_of_unique_emails)
    SendEmails()

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
        print("waiting to delete temporary files (./temp/*) (pdfs will be safe) (don't close please)")
        time.sleep(5)
        if os.path.exists("files.db"):
            os.remove("files.db")
        if os.path.exists("files.db-journal"):
            os.remove("files.db-journal")

