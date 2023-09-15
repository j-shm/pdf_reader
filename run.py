from pdf2image import convert_from_path
import pytesseract
import win32com.client
import datetime
import glob
import os
import sqlite3

#options 

company_name = "company"
send_from = ""

one_at_a_time = True
#end options


#debug options

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

def GetImageFromPdf(pdf):
    return convert_from_path(pdf,poppler_path=os.getcwd()+"/poppler/Library/bin")[0]

#make singular function with diff paramaters!

def CropName(pdf_image):
    x,y = pdf_image.size
    cropped_pdf_image = pdf_image.crop((530,240,990,261))
    return cropped_pdf_image

def CropEmailAddress(pdf_image):
    x,y = pdf_image.size
    cropped_pdf_image = pdf_image.crop((271.8,577.5,979.3,603.6))
    return cropped_pdf_image

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

    for account in ol.Session.Accounts:
        if account.DisplayName == send_from:
            print("email found.")
            print(str(account) + "=" + send_from)
            newmail._oleobj_.Invoke(*(64209, 0, 8, 0, account))

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

def Closer():
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

if __name__ == '__main__':
    company_name = input("Enter company name: ")
    print("COMPANY NAME: " + company_name)
    send_from = input("Enter email to send from: ")
    print("EMAIL: " + send_from)
    print("")
    selection = input("Type y to continue:")
    if selection != "y":
        exit()

    errors = ""
    pdfs = GetPdf()
    for pdf in pdfs:
        pdf_img = GetImageFromPdf(pdf)

        email_img = CropEmailAddress(pdf_img)
        name_img = CropName(pdf_img)

        print("reading images...")
        name = pytesseract.image_to_string(name_img).strip()
        email = pytesseract.image_to_string(email_img).strip()
        print("read images!")
        
        cur.execute(f'INSERT INTO FILES(pdf,email,name) VALUES(?,?,?)',(pdf,email,name))


    #list_of_unique_emails = cur.execute("SELECT DISTINCT email,name FROM FILES").fetchall()

    SendEmails()