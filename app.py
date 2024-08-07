import datetime
from fasthtml.common import *
import pdfplumber
import glob

import win32com
import win32com.client

app,rt,files,Email = fast_app(
    'data/files.db',
    id=int, email=str, name=str, pdf=str, pk='id')



def send_handler(companyname:str,emailname:str):
    email_to_send = files()[0]
    if companyname == "" or emailname == "":
        return "Please enter a company name and email"
    if email_to_send:
        outlook_email(email_to_send.name, email_to_send.email, GetExcel(email_to_send.pdf), companyname,emailname)
        files.delete(email_to_send.id)
    else:
        return "No emails to send"
    return GetTable() + f"<h1>Last Opened: {email_to_send.pdf}</h1>"
    


def GetExcel(name):
    tname = name.split(".")[0]
    excel_files_1 = glob.glob("pdf/*.xsls")
    excel_files_2 = glob.glob("pdf/*.xlsx")
    excel_files = excel_files_1 + excel_files_2
    matches = [string for string in excel_files if tname in string]
    return matches


def outlook_email(name, email, attachments, company_name,email_account):
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'{company_name} - {GetDate()} - {name.split(".")[0]}'
    newmail.To = '' + email

    for attach in attachments:
        newmail.Attachments.Add(os.getcwd() + "\\" + attach)

    for account in ol.Session.Accounts:
        if account.DisplayName == email_account:
            newmail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
            newmail.Display()
    newmail.Display()

def GetDate():
    now = datetime.datetime.now()
    month = now.strftime('%B')
    year = now.year
    return f'{month} {year}'

def GetPdf():
    pdfs = []
    for file in glob.glob("pdf/*.pdf"):
        pdfs.append(file)
    return pdfs

def PdfExists(pdf):
    for file in files():
        if file.pdf == pdf:
            return True
    return False

@app.route("/table", methods=['post'])
def GetTable():
    print("Getting table")
    if(len(files()) == 0):
        return "<h1>No emails left 🐰🎊🎊🎊🎊</h1>"
    table = f"<h1>{len(files())} emails left</h1><table><tr><th>Name</th><th>Email</th><th>PDF</th></tr>"
    for file in (files()):
        table += "<tr><td>" + file.name + "</td><td>" + file.email + "</td><td>" + file.pdf + "</td></tr>"
    return table

@app.route("/", methods=['post'])
def process_pdfs():
    pdfs = GetPdf()
    for pdf in pdfs:
        if PdfExists(pdf):
            continue
        emails = []
        names = []
        email = ""
        name = ""
        with pdfplumber.open(pdf) as plumber_pdf:
            first_page = plumber_pdf.pages[0]
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
        if len(names[0]) > 0:
            name = names[0]
        if len(emails) > 0:
            for listed_email in emails:
                email += listed_email + "; "
        files.insert(Email(email=email, name=name, pdf=str(pdf)))
    return GetTable()

@app.route("/delete", methods=['post'])
def delete():
    print(files)
    for file in files():
        files.delete(file.id)
    return GetTable()

@app.route("/send", methods=['post'])
def send_email(companyname:str,emailname:str):
    return send_handler(companyname,emailname)
    

@app.get("/")
def get():
    add = Form(
        Group(Button("Get Emails")), hx_post="/", target_id='gen-list', hx_swap="innerHTML"
        )
    email_label = H2('💌')
    company_label = H2('🏢')
    company = Input(id="new-company-name", name="companyname", placeholder="Enter the company name")
    email = Input(id="new-email", name="emailname", placeholder="Enter your email")
    send_email = Form(
        Group(company_label,company,email_label,email,Button("Send Email")), hx_post="/send", target_id='gen-list', hx_swap="innerHTML"
    )
    breaker = Br(Br(Br(Br(Br(Br(Br(Br(Br(Br())))))))))
    gen_list = Div(id='gen-list')
    test = Form(Group(Button("Refresh")), hx_post="/table", target_id='gen-list', hx_swap="innerHTML") 
    delete = Form(Group(Button("Delete")), hx_post="/delete", target_id='gen-list', hx_swap="innerHTML")  
    
    return Title('Email sender 🐰'), Main(H1('Email sender 🐰'), send_email,gen_list,breaker,add,test,delete, cls='container')


    


serve()