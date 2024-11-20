import smtplib
import ssl
from email.message import EmailMessage
import pandas as pd
from datetime import datetime
from tqdm.auto import tqdm

# Define email sender credentials
email_sender = 'test@gmail.com'
email_password = '*****' # got from https://myaccount.google.com/u/2/apppasswords

# path of the Excel file
file_path = 'log_file.xlsx'

# reading the Excel file
df = pd.read_excel(file_path)

# selecting the rows where the data is NaN
emails_to_send = df[df['DATA INVIO'].isna()]

# taking the emails
emails_receiver = emails_to_send['EMAIL'].tolist()

# subject of the email
subject = 'Subject test'
# body of the email
body = """
Body test
of the email.

"""

# adding SSL (layer of security)
context = ssl.create_default_context()

# log in and send the email
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)
    for rec in tqdm(emails_receiver):
        em = EmailMessage()
        em['From'] = email_sender
        em['Subject'] = subject
        em['To'] = rec
        em.set_content(body)
        smtp.sendmail(email_sender, rec, em.as_string())

        # updating the column "DATA INVIO" for the corresponding row with the current date
        row_index = df[df['EMAIL'] == rec].index[0]
        df.at[row_index, 'DATA INVIO'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # update the file_path with the updated DataFrame
        df.to_excel(file_path, index=False)
