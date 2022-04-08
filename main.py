import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = p.read_excel("students.xlsx")
print(type(data))
email_col = data.get("email")

list_of_emails = list(email_col)
print(list_of_emails)

try:
    server = sm.SMTP("smtp.gmail.com", 465)
    server.starttls()
    server.login("gmail","password" )
    from_ = ""
    to_ =list_of_emails
    message = MIMEMultipart("alternative")
    message['subject']="This is just Testing Message"
    message["from"]="gmail"

    html='''
    <html>
    <head>
    
    </head>
    <body>
    <h1>learn code with pratik</h1>
    <h2>hello Everyone</h2>
    <p>this is just testing para</p>
    <button style="padding:20px;background:green;color:white;">Varify</button>
    
    </body>
    </html>
    
    '''
    text = MIMEText(html, "html")
    message.attach(text)

    server.sendmail(from_, to_, message.as_string())
    print("message has been sent to the emails.")
        # less secure app search and turn on
except Exception as e:
    print(e)
    print(e)