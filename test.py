import win32com.client

path_message_file = r'C:\Users\Edison\Diajo-SAC\python\message\message.html' #the path will be changeable

outlook = win32com.client.Dispatch('Outlook.Application')
email_content = open(path_message_file, "r", encoding="utf-8").read()


account_send = None

for acc_send in outlook.Session.Accounts:
    if acc_send.SmtpAddress == "diajomkt@diajosac.com":
        account_send = acc_send

count_name = 0

mail = outlook.CreateItem(0)
mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_send))
mail.To = "pyamakahua@esan.edu.pe"
mail.Subject = "Hola Peter! somos Diajo SAC"
mail.HTMLBody = email_content
mail.Send()