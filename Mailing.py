import pandas as pd
import numpy as np
import smtplib
from email.message import EmailMessage
import json
from string import Template

def Send_Mails(server, subject, data, Mail_Format):
    # Load sender email from configuration file
    with open('Data/Sender-config.json', 'r') as config_file:
        config = json.load(config_file)
        Sender_Email = config['Sender-mail-id']

    # Iterate through each row in the DataFrame
    for index, row in data.iterrows():
        # Prepare a dictionary of variable names and values for the current row
        variables = {}
        for var_name in data.columns:
            # Clean the variable name by replacing spaces, dashes, colons, and dots with underscores
            cleaned_var_name = var_name.replace(' ', '_').replace('-', '_').replace(':', '').replace('.', '')
            value = row[var_name]

            # Handle NaN values
            if pd.isna(value):
                value = "NaN"

            variables[cleaned_var_name] = value

        # Format the email content with variables
        formatted_mail = Mail_Format.format(**variables)

        # Send the email using post_mail function (assuming it exists and takes appropriate arguments)
        post_mail(server, subject, Sender_Email, row['Email'], formatted_mail)

def post_mail(server, Subject, Sender_Email, Recipient_Email, Mail_Format):
    
    msg = EmailMessage()
    
    msg['From'] = Sender_Email

    msg['To'] = Recipient_Email

    msg['Subject'] = Subject
    
    msg.set_content(Mail_Format)

    try:

        server.send_message(msg)
        
        print(f"Email sent successfully to {Recipient_Email}!")
    
    except smtplib.SMTPAuthenticationError:
        # This exception is raised for authentication errors
        print("Failed to authenticate. Check your email and password.")
    except smtplib.SMTPException as e:
        # This exception is raised for other SMTP errors
        print("SMTP error occurred:", e)
    except Exception as e:
        # This catches any other exceptions that might occur
        print(f"Failed to send email to {Recipient_Email}:", e)
        
