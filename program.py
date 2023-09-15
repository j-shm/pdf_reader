import os
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path
import pytesseract
import win32com.client
import datetime
import glob

# Initialize Tkinter
root = tk.Tk()
root.title("PDF Emailer")
root.geometry("400x400")


# Options
company_name = "Company"
send_from = tk.StringVar()
one_at_a_time = True

# Debug options
delete_db = True




# Create SQLite database and cursor
if os.path.exists("files.db"):
    os.remove("files.db")
if os.path.exists("files.db-journal"):
    os.remove("files.db-journal")

con = sqlite3.connect("files.db")
cur = con.cursor()

table = """
CREATE TABLE IF NOT EXISTS FILES (
    pdf text PRIMARY KEY,
    email text,
    name text
);
"""
cur.execute(table)


def select_directory():
    status_text.config(text="Processing PDFs..." )
    directory = filedialog.askdirectory()
    if directory:
        os.chdir(directory)
        pdfs = GetPdf()
        process_pdfs(pdfs)


def process_pdfs(pdfs):
    errors = ""
    main_path = os.getcwd() + "\\temp_img\\"
    for pdf in pdfs:
        
        pdf_img = GetImageFromPdf(pdf)

        email_img = CropEmailAddress(pdf_img)
        name_img = CropName(pdf_img)

        print("reading pdf " + pdf + "...")

        name = pytesseract.image_to_string(name_img).strip()
        email = pytesseract.image_to_string(email_img).strip()

        print("read pdf " + pdf + "...")

        path = main_path + pdf.split(".")[0] +"\\"
        if not os.path.exists(path):
            os.makedirs(path)
            pdf_img.save(f"{path}pdf.png")
            email_img.save(f"{path}email.png")
            name_img.save(f"{path}name.png")
        

        cur.execute('INSERT INTO FILES(pdf,email,name) VALUES(?,?,?)', (pdf, email, name))

    status_text.config(text=f"Processed {len(pdf)} pdfs")
    print("Temp files saved to " + main_path)
    send_emails_one_at_a_time_ui()

file_iterator = None

def send_emails_one_at_a_time_ui():
    items = cur.execute("SELECT * FROM FILES").fetchall()
    global file_iterator
    file_iterator = iter(items)


def next_email():
    try:
        item = next(file_iterator)
        details_text ="Details:\n" + "From: " + company_name + "\n" + item[0] + "\n" +item[1] + "\n" + item[2]
        email_details_label.config(text=details_text)
        open_email(item)
    except StopIteration:
        email_details_label.config(text="All emails done!")
        Closer()

def open_email(item):
        details_text ="Details:\n" + "From: " + company_name + "\n" + item[0] + "\n" +item[1] + "\n" + item[2]
        email_details_label.config(text=details_text)
        attachments = [item[0]]
        email = item[1]
        name = item[2]
        excels = GetExcel(item[0].split(".")[0])
        for excel in excels:
            if excel not in attachments:
                attachments.append(excel)
        send_email(name, email, attachments)

def send_emails_one_at_a_time():
    items = cur.execute("SELECT * FROM FILES").fetchall()
    for item in items:
        details_text ="Details:\n" + "From: " + company_name + "\n" + item[0] + "\n" +item[1] + "\n" + item[2]
        email_details_label.config(text=details_text)
        input()
        attachments = [item[0]]
        email = item[1]
        name = item[2]
        excels = GetExcel(item[0].split(".")[0])
        for excel in excels:
            if excel not in attachments:
                attachments.append(excel)
        send_email(name, email, attachments)



def send_emails():
    list_of_unique_emails = cur.execute("SELECT DISTINCT email,name FROM FILES").fetchall()
    for tup_email in list_of_unique_emails:
        email = tup_email[0]
        name = tup_email[1]

        attachments = []

        sql_attachments = cur.execute("SELECT pdf FROM FILES WHERE email = ? AND name = ?", (email, name), ).fetchall()
        for attachment in sql_attachments:
            attachments.append(attachment[0])
            excels = GetExcel(attachment[0].split(".")[0])
            for excel in excels:
                if excel not in attachments:
                    attachments.append(excel)
        send_email(name, email, attachments)


def send_email(name, email, attachments):
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = f'{company_name} - {GetDate()} - {name.split(".")[0]}'
    newmail.To = '' + email

    for attach in attachments:
        newmail.Attachments.Add(os.getcwd() + "\\" + attach)

    for account in ol.Session.Accounts:
        if account.DisplayName == send_from.get():
            newmail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
            newmail.Display()


def GetImageFromPdf(pdf):
    return convert_from_path(pdf, poppler_path=os.getcwd() + "/poppler/Library/bin")[0]


def CropName(pdf_image):
    x, y = pdf_image.size
    cropped_pdf_image = pdf_image.crop((530, 240, 990, 261))
    return cropped_pdf_image


def CropEmailAddress(pdf_image):
    x, y = pdf_image.size
    cropped_pdf_image = pdf_image.crop((271.8, 577.5, 979.3, 603.6))
    return cropped_pdf_image


def GetDate():
    now = datetime.datetime.now()
    month = now.strftime('%B')
    year = now.year
    return f'{month} {year}'


def GetPdf():
    pdfs = []
    for file in glob.glob("*.pdf"):
        pdfs.append(file)
    return pdfs


def GetExcel(name):
    tname = name.split(".")[0]
    excel_files_1 = glob.glob("*.xsls")
    excel_files_2 = glob.glob("*.xlsx")
    excel_files = excel_files_1 + excel_files_2
    matches = [string for string in excel_files if tname in string]
    return matches


def Closer():
    if errors != "":
        print("ERRORS:")
        print(errors)
    con.commit()
    cur.close()
    con.close()
    if delete_db:
        import time
        print("Waiting to delete temporary files (./temp/*) (PDFs will be safe) (please don't close)")
        time.sleep(5)
        if os.path.exists("files.db"):
            os.remove("files.db")
        if os.path.exists("files.db-journal"):
            os.remove("files.db-journal")


if __name__ == '__main__':
    print("Please use the ui")
    label = tk.Label(root, text="PDF Emailer", font=("Arial", 16))
    label.pack(pady=20)

    ol = win32com.client.Dispatch("outlook.application")

    accounts = []
    for account in ol.Session.Accounts:
        accounts.append(account.DisplayName)
    ol.Quit()


    company_name_label = tk.Label(root, text="Enter company name:")
    company_name_label.pack()

    company_name_entry = tk.Entry(root)
    company_name_entry.pack(pady = 10)

    send_from_label = tk.Label(root, text="Enter email to send from:")
    send_from_label.pack()

    send_from.set(accounts[0])
    dropdown = tk.OptionMenu(root, send_from, *accounts)
    dropdown.pack(pady = 10)

    select_button = tk.Button(root, text="Select Directory", command=select_directory)
    select_button.pack(pady = 20)

    status_text = tk.Label(root, text="Waiting for directory...")
    status_text.pack()

    email_details_label = tk.Label(root, text="", font=("Arial", 12))
    email_details_label.pack(pady=10)
    
    next_item_button = tk.Button(root, text="Next Item", command=next_email)
    next_item_button.pack(pady = 20)


    root.mainloop()