import pandas as pd
import numpy as np
import smtplib
from email.message import EmailMessage
import json
from string import Template
import tkinter as tk
from tkinter import ttk, messagebox
import datetime

def Send_Mails(server, subject, data, Mail_Format):
    # Load sender email from configuration file
    with open('Data/Sender-config.json', 'r') as config_file:
        config = json.load(config_file)
        Sender_Email = config['Sender-mail-id']

    # creating console logs popup
    Logs_PopUp()

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
        
    messagebox.showinfo("Information", "ALl mails are sented")
    PushToConsoleLogs("---------- All Mails Are Sented! ----------")
# Posting emails
def post_mail(server, Subject, Sender_Email, Recipient_Email, Mail_Format):
    
    msg = EmailMessage()

    # from
    msg['From'] = Sender_Email

    # to
    msg['To'] = Recipient_Email

    # subject
    msg['Subject'] = Subject

    # content
    msg.set_content(Mail_Format)

    try:

        server.send_message(msg)
        
        print(f"Email sent successfully to {Recipient_Email}!")
        # pushing to console
        PushToConsoleLogs(f"Email sent successfully to {Recipient_Email}!")
    
    except smtplib.SMTPAuthenticationError:
        # This exception is raised for authentication errors
        print("Failed to authenticate. Check your email and password.")
        messagebox.showerror("Error", "Failed to authenticate. Check your email and password.")
    except smtplib.SMTPException as e:
        # This exception is raised for other SMTP errors
        print("SMTP error occurred:", e)
        messagebox.showerror("Error", f"SMTP error occurred:{e}")
    except Exception as e:
        # This catches any other exceptions that might occur
        print(f"Failed to send email to {Recipient_Email}:", e)
        messagebox.showerror("Error", f"Failed to send email to {Recipient_Email}:{e}")


def Logs_PopUp():
    global text_widget

    window = tk.Tk()
    window.geometry('400x300')
    window.title("Console Logs")
    
    # Title label
    title = tk.Label(window, text="Console Logs Session", font=("Arial", 14, "bold"))
    title.pack(padx=10, pady=10)
    
    # Text widget for logs
    text_widget = tk.Text(window, wrap='word', width=100, height=15, state=tk.DISABLED)
    text_widget.pack(padx=10, pady=10)

    # Button to save logs
    save_button = ttk.Button(window, text="Save Logs", command=lambda: save_logs_to_file(text_widget))
    save_button.pack(pady=10)

# saving logs as text file
def save_logs_to_file(text_widget):
    logs = text_widget.get("1.0", tk.END)
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"Logs/console_logs_{current_datetime}.txt"
    try:
        with open(filename, 'w') as file:
            file.write(logs)
        messagebox.showinfo("Logs Saved", f"Logs successfully saved to {filename}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save logs: {e}")

# pushing logs to console
def PushToConsoleLogs(log):
    text_widget.config(state=tk.NORMAL)
    text_widget.insert(tk.END, f"{log}\n")
    text_widget.config(state=tk.DISABLED)

