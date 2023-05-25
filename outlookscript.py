import os
import subprocess
import sys
import win32com.client

def open_outlook_with_file_path(file_path):
    # Extract the file name from the file path
    file_name = os.path.basename(file_path)

    # Open Outlook and create a new email
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail_item = outlook.CreateItem(0)
    mail_item.Subject = f"[FOR REVIEW] {file_name}"

    #Format the file as a hyperlink and paste it in the email
    file_link = f"<a href='file://{file_path}'>{file_path}</a>"
    mail_item.HTMLBody = f"File located here:<br/><br/>{file_link}"

    # Display the email
    mail_item.Display()

def main():
    # Check if a file path is provided as an argument
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        open_outlook_with_file_path(file_path)
    else:
        print("No file path provided!")

if __name__ == "__main__":
    main()
