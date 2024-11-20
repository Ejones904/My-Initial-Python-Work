# -*- coding: utf-8 -*-

import pandas as pd
import smtplib
from email.mime.text import MIMEText

# Email template stored as a multi-line string
email_template = """

Dear {Contact_Name},

I wanted to follow with you on Salesforce Case {Case_Number} that was opened on {Opened_Date} about "{Subject}". Has this issue been resolved?  

Thank you ,


"""

# Load the Excel file
file_path =   # Replace with your Excel file path
data = pd.read_excel(file_path,header=0)


# Email credentials
sender_email =   # Replace with your email
password =   # Replace with your email password

# Set up SMTP connection
smtp_server = 'smtp.gmail.com'  # Replace with your provider's SMTP server (e.g., Gmail)
smtp_port = 587  # Typically 587 for TLS

try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # Upgrade the connection to a secure encrypted connection
    server.login(sender_email, password)
   
    for index, row in data.iterrows():
        # Format the email content
        email_content = email_template.format(
            Case_Number=row['Case_Number'],
            Contact_Name=row['Contact_Name'],
            Opened_Date=row['Opened_Date'],
            Subject=row['Subject']
        )
        # Create the email
        msg = MIMEText(email_content, 'plain')
        msg['From'] = sender_email
        msg['To'] = row['Contact_Email']
        msg['Subject'] = f"Trimble Support Case {row['Case_Number']} Follow-Up: {row['Subject']}"
        
        # Send the email
        server.sendmail(sender_email, row['Contact_Email'], msg.as_string())
        print(f"Email sent to {row['Contact_Email']}")
except Exception as e:
    print(f"Failed to send emails: {e}")
finally:
    server.quit()  
       