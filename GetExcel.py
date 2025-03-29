import requests
from email.message import EmailMessage
import ssl
import smtplib
from dotenv import load_dotenv
import os

load_dotenv()  # Load from .env

email_sender = os.getenv("EMAIL_SENDER")
email_password = os.getenv("EMAIL_PASSWORD")
outlock_mail = os.getenv("OUTLOOK_EMAIL")
outlock_pass = os.getenv("OUTLOOK_PASS")
access_token = os.getenv("ACCESS_TOKEN")


use_ssl=False
headers = {
    'Authorization': f'Bearer {access_token}'
}


url = "https://graph.microsoft.com/v1.0/me/drive/root:/Livro1.xlsx:/workbook/worksheets('Folha1')/usedRange"

response = requests.get(url, headers=headers)
data = response.json()
listMails =[]
for row in data.get("values", []):
    listMails.append(str(row[0]))
print(data)
subject = "Excited to Share Something New With You!"
body = """
Hi there,

I just finished working on a new project and wanted to share it with you — it's called 'Boot'. 
Would love to hear what you think.

Let me know if you're curious, and I’ll send you more details.

Best,  
Simao
"""
microsoft_domains = [
    "outlook.com", "outlook.pt", "outlook.es", "outlook.fr", "outlook.it", "outlook.de", "outlook.com.br", "outlook.co.uk",
    "hotmail.com", "hotmail.pt", "hotmail.es", "hotmail.fr", "hotmail.it", "hotmail.de", "hotmail.co.uk",
    "live.com", "live.pt", "live.fr", "live.it", "live.com.pt", "live.ca", "live.co.uk",
    "msn.com"
]

for mail in listMails:
    print(f"\nSending to: {mail}")
    domain = mail.split('@')[1].lower()

    em = EmailMessage()
    em['To'] = mail
    em['Subject'] = subject
    em.set_content(body)

    context = ssl.create_default_context()

    if domain in microsoft_domains:
        smtp_server = "smtp.office365.com"
        smtp_port = 587
        sender_email = outlock_mail
        sender_pass = outlock_pass
        em['From'] = sender_email
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls(context=context)
            smtp.login(sender_email, sender_pass)
            smtp.sendmail(sender_email, mail, em.as_string())
    else:
        smtp_server = 'smtp.gmail.com'
        smtp_port = 465
        sender_email = email_sender
        sender_pass = email_password
        em['From'] = sender_email
        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as smtp:
            smtp.login(sender_email, sender_pass)
            smtp.sendmail(sender_email, mail, em.as_string())

    print(f"✔️ Sent successfully using {smtp_server}")