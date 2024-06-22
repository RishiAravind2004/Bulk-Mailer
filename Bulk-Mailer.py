"""
        Bulk Mailer System 

    - to send bulk of mails from excel with an format of mail with its concurrent values

-by Rishi Aravaind

"""

# Module Imports
import json
import smtplib
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from Mailing import Send_Mails

# Define global variables
current_file_path = ""
selected_file_label = None
df = None


"""

            Verify Session
    
"""

# Verify session function
def Verify_Session(window):

    # Title lable
    title = tk.Label(window, text="Welcome to Bulk Mailer!\n\nVerfication Session", font=("Arial", 14, "bold"))
    title.pack(padx=10, pady=10)
    
    text = tk.Label(window, text="Please, verify your sender gmail authentication configuration. If you don't know please checkout the project documention for setup!\nDocumentation link: https://github.com/RishiAravind2004/Bulk-Mailer")
    text.pack(padx=10, pady=10)
    
    VerifyBtn = tk.Button(window, text="Verify", command=CheckSenderConfig)
    VerifyBtn.pack(padx=10, pady=10)

# Read configuration from JSON file and process to validation
def CheckSenderConfig():
    try:
        with open('Data/Sender-config.json', 'r') as file:
            config = json.load(file)

        # Access configuration settings
        app_password = config['App-password']
        sender_mail_id = config['Sender-mail-id']

        print("Configuration key found!")
        
        # Validate the credentials
        ValidateSenderMail(sender_mail_id, app_password)

    except FileNotFoundError:
        print("Error: JSON file not found.")
        messagebox.showerror("Error", "JSON file not found.")
    except json.JSONDecodeError as e:
        print(f"Error reading JSON file: {e}")
        messagebox.showerror("Error", "Error reading JSON file.")
    except KeyError as e:
        print(f"Configuration key not found: {e}")
        messagebox.showerror("Error", "Configuration key not found.")

# Function for checking sender credentials are valid
def ValidateSenderMail(sender_mail_id, app_password):
    global server
    
    try:
        # Trying to login into smtp server with email and app password
        server = smtplib.SMTP('smtp.gmail.com', 587)
        
        # Start TLS (Transport Layer Security) for secure communication
        server.starttls()

        # Login to the SMTP server with your email and app-specific password
        server.login(sender_mail_id, app_password)
        messagebox.showinfo("Information", "Credentials Validated!")
        print("Sender's Gmail credentials have been validated!")

        # If validated, move to further process
        Selecting_file_session()
        
    except Exception as e:
        # Can't log in to the SMTP server
        print("Sender's Gmail credentials are not valid!")
        print(e)
        messagebox.showerror("Error", "Please check the sender mail credentials in the configuration file.")

"""

            File Session
    
"""

# Selecting file & processing session
def Selecting_file_session():
    global selected_file_label

    # Purge current widget
    destroy_current_session_content(window)

    window.geometry('800x300')

    # Title lable
    title = tk.Label(window, text="File Selecting Session", font=("Arial", 14, "bold"))
    title.pack(padx=10, pady=10)
    
    text = tk.Label(window, text="Select your excel file for Recipient datas and Email format for sending mails")
    text.pack(padx=10, pady=10)

    selected_file_label = tk.Label(window, text="No file selected")
    selected_file_label.pack(padx=10, pady=10)

    open_button = tk.Button(window, text="Open Excel File", command=open_file_dialog)
    open_button.pack(padx=10, pady=10)

    ProcessBtn = tk.Button(window, text="Process File", command=lambda: process_selected_file() if current_file_path else messagebox.showerror("Error", "Select a file before processing."))
    ProcessBtn.pack(padx=10, pady=10)

# dialog box for choosing excel file
def open_file_dialog():
    global current_file_path

    # get current selected excel file path
    current_file_path = filedialog.askopenfilename(title="Select an Excel File", filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
    
    if current_file_path:
        selected_file_label.config(text=f"Selected File: {current_file_path}")

# just processing
def process_selected_file():
    try:
        check_for_mail_format(current_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# checking for mail format file
def check_for_mail_format(excel_file_path):
    try:
        with open('Data/Mail-format.txt', 'r') as file:
            mail_format_lines = file.readlines()

        if len(mail_format_lines) > 3: # checking the mail format lines is above 3
            process_file(excel_file_path) # processing the file
        else:
            print("Error: The Mail format file has 3 or fewer lines.")
            messagebox.showerror("Error", "The Mail format file has 3 or fewer lines.")
    except FileNotFoundError:
        print("Error: The Mail format file does not exist.")
        messagebox.showerror("Error", "The Mail format file does not exist.")
    except Exception as e:
        print(f"Error: An error occurred: {e}.")
        messagebox.showerror("Error", f"An error occurred: {e}.")

# processing the and check the excel file is valid and "Email" column is exist
def process_file(file_path):
    global df
    try:
        df = pd.read_excel(file_path) # excel file into DataFrames
        if 'Email' not in df.columns: # checking wheather the dataframe having "Email" column, coz using this only we going to send mails
            messagebox.showerror("Error", "No 'Email' column found in the Excel file.")
        else:
            Display_Session(df) # moving to display the excel file content
    except FileNotFoundError:
        selected_file_label.config(text="Error: The selected Excel file does not exist.")
    except Exception as e:
        selected_file_label.config(text=f"Error: {str(e)}")


"""

            Display & Editing Session
    
"""

# Display Session function
def Display_Session(df):
    # adjust the window root size
    window.geometry('800x600')
    
    global treeview, save_status_label
    
    destroy_current_session_content(window)

    # Title lable
    title = tk.Label(window, text="Recipient File Datas", font=("Arial", 14, "bold"))
    title.pack(padx=10, pady=10)

    # Notes
    Notes = tk.Label(window, text="Notes: You may can edit here itself and make your changes saves maded!")
    Notes.pack(padx=10, pady=10)

    # Display excel file content
    display_table(df)

    # Label to display save status
    save_status_label = tk.Label(window, text="")
    save_status_label.pack(padx=10, pady=10)
    
    # Button to save the file
    saveBtn = tk.Button(window, text="Save File", command=save_file)
    saveBtn.pack(padx=10, pady=10)

    # Button to next display session
    nextBtn = tk.Button(window, text="Next",  command=Display_Session_2)
    nextBtn.pack(padx=10, pady=10)

# --------------------------------------------------------------------------- #

# Function to display DataFrame as a table in tkinter
def display_table(df):
    global treeview
    
    # Create Treeview widget for displaying table
    treeview = EditableTreeview(window, show='headings')
    treeview.pack(padx=20, pady=20)

    # Insert DataFrame columns as table columns
    treeview["columns"] = list(df.columns)
    for col in df.columns:
        treeview.column(col, anchor="w", width=100)
        treeview.heading(col, text=col)
    
    # Insert DataFrame rows as table rows
    for index, row in df.iterrows():
        treeview.insert("", tk.END, values=list(row), iid=index)

    # Limit the number of rows displayed in the table view
    treeview.config(height=min(20, len(df)))

# Function for editing cell in tree view
class EditableTreeview(ttk.Treeview):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._entry = None
        self.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        region = self.identify("region", event.x, event.y)
        if region == "cell":
            column = self.identify_column(event.x)
            row = self.identify_row(event.y)
            if self._entry:
                self._entry.destroy()
            x, y, width, height = self.bbox(row, column)
            value = self.item(row, "values")[int(column[1:]) - 1]
            self._entry = tk.Entry(self, width=width)
            self._entry.place(x=x, y=y, width=width, height=height)
            self._entry.insert(0, value)
            self._entry.focus()
            self._entry.bind("<Return>", lambda e: self.save_edit(row, column))
            self._entry.bind("<FocusOut>", lambda e: self.save_edit(row, column))

    def save_edit(self, row, column):
        value = self._entry.get()
        col_index = int(column[1:]) - 1
        self.set(row, column=col_index, value=value)
        self._entry.destroy()
        self._entry = None

        # Update DataFrame with explicit type conversion using .iloc
        dtype = df.dtypes.iloc[col_index]
        if dtype == "float64":
            df.iloc[int(row), col_index] = float(value)
        elif dtype == "int64":
            df.iloc[int(row), col_index] = int(value)
        else:
            df.iloc[int(row), col_index] = value

# Function to save the DataFrame to the Excel file
def save_file():
    global save_status_label
    if current_file_path:
        try:
            df.to_excel(current_file_path, index=False)
            save_status_label.config(text="File saved successfully!")
        except Exception as e:
            save_status_label.config(text=f"Error saving file: {str(e)}")
    else:
        save_status_label.config(text="No file selected to save.")

# --------------------------------------------------------------------------- #


# Second part of display session
def Display_Session_2():
    
    destroy_current_session_content(window)

    global save_status_label, Mail_Format

    # Title
    title = tk.Label(window, text="Email Format", font=("Arial", 14, "bold"))
    title.pack(padx=10, pady=10)

    # Notes
    Notes = tk.Label(window, text="Notes: You may can edit here itself and make your changes saves maded!")
    Notes.pack(padx=10, pady=10)
    
    # Display the format of email
    display_mail_format()

    # Label to display save status
    save_status_label = tk.Label(window, text="")
    save_status_label.pack(padx=10, pady=10)
    
    # Button to save the file
    saveBtn = tk.Button(window, text="Save File", command=save_text_file)
    saveBtn.pack(padx=10, pady=10)
    
    # Button to previous display session
    nextBtn = tk.Button(window, text="Back",  command= lambda: process_file(current_file_path))
    nextBtn.pack(padx=10, pady=10)

    # Button to send mails session
    SendMailsBtn = tk.Button(window, text="Send Mails",  command= lambda: Send_Mails(server, get_subject(), df, Mail_Format))
    SendMailsBtn.pack(padx=10, pady=10)

# --------------------------------------------------------------------------- #

# Function to display text file content in a Text widget in tkinter
def display_mail_format():
    global text_widget, save_status_label, Mail_Format
    
    with open('Data/Mail-format.txt', 'r') as file:
            mail_format_content = file.read()
            Mail_Format = mail_format_content
    
    text_widget = tk.Text(window, wrap='word', width=100, height=25)
    text_widget.pack(padx=20, pady=20)
    text_widget.insert("1.0", mail_format_content)

# Function to save the edited text content to the Text file
def save_text_file():
    global save_status_label
    if 'Data/Mail-format.txt':
        try:
            with open('Data/Mail-format.txt', 'w') as file:
                file.write(text_widget.get("1.0", tk.END))
            save_status_label.config(text="File saved successfully!")
        except Exception as e:
            save_status_label.config(text=f"Error saving file: {str(e)}")
    else:
        save_status_label.config(text="No file selected to save.")

# --------------------------------------------------------------------------- #

"""

            Function for operations
    
"""

# Function to destroy current contents on window
def destroy_current_session_content(window):
    for widget in window.winfo_children():
        widget.destroy()

# Pop-up for getting subject for email
def get_subject():
    root = tk.Tk()
    root.withdraw()
    subject = simpledialog.askstring("Email Subject", "Please enter the subject of the email:")
    root.destroy()
    return subject

"""

            Starting of program
    
"""

# Starts
if __name__ == "__main__":
    window = tk.Tk()
    window.geometry('800x200')
    window.title("Bulk Mailer")
    img = PhotoImage(file='images/Logo.png')
    window.iconphoto(True, img)
    window.iconbitmap('images/Logo.ico')
    Verify_Session(window)
    window.mainloop()
