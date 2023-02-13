import win32com.client
import pandas as pd 
import openpyxl
import datetime
import re

path_excel_file = r'C:\Users\Edison\Diajo-SAC\python\file-db\Filtrado BASE GRANDES Y MEDIANAS EMPRESAS - PERÚ.xlsx' #the path will be changeable
path_message_file = r'C:\Users\Edison\Diajo-SAC\python\message\message.html' #the path will be changeable

data_pd_del = pd.read_excel(path_excel_file)

outlook = win32com.client.Dispatch('outlook.application')

mapi = outlook.GetNamespace("MAPI")

account_del = mapi.Folders("diajomkt@diajosac.com")

inbox = account_del.Folders("Bandeja de entrada")

messages = inbox.Items

today = datetime.datetime.now().date()

messages = messages.Restrict("[Subject] = 'Cancelar suscripción'")


#Deleteing ------------------------------------------------------

workbook = openpyxl.load_workbook(path_excel_file)
sheet = workbook.active
sheet = workbook["GRAN Y MEDIANA EMPRESA - PERÚ"]

account_to_del = None

for msg in messages:
    # Change the frequency of reading emails according to the frequency of sending advertising
    if msg.ReceivedTime.date() == today:
        if msg.SenderEmailType == "EX":
            account_to_del = msg.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            account_to_del = msg.SenderEmailAddress

        for row in sheet.rows:
            if row[10].value == account_to_del: #changeable column number
                row_del = row[10].row
                sheet.delete_rows(row_del, 1)
                workbook.save(path_excel_file)
                break


#Send emails ----------------------------------------------------

data_pd = pd.read_excel(path_excel_file, header=0)

email_content = open(path_message_file, "r", encoding="utf-8").read()

r_emails = data_pd['E-MAIL']
r_names = data_pd['NOMBRE / REPRESENTANTE LEGAL']

account_send = None

for acc_send in outlook.Session.Accounts:
    if acc_send.SmtpAddress == "diajomkt@diajosac.com":
        account_send = acc_send

count_name = 0

errors = []

for email in r_emails:
    try:
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_send))
        mail.To = email
        name = re.findall(r'\b\w+', r_names[count_name])
        if name: 
            mail.Subject = "Hola " + name[-1].title() + "!, somos Diajo SAC"
            count_name += 1
        else:
            mail.Subject = "Hola! somos Diajo SAC"
            count_name += 1
        mail.HTMLBody = email_content
        mail.Send()
    except ValueError as e:
        errors.append(e)