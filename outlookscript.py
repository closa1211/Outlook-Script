import os
import subprocess
import sys

def open_outlook_with_file_path(file_path):
    # Format the file path for the email body
    email_body = f"File located here: {file_path}"

    # Open Outlook with the email body
    subprocess.run(['start', 'outlook', '/c', 'ipm.note', '/m', email_body], shell=True)

def main():
    # Check if a file path is provided as an argument
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if os.path.isfile(file_path):
            open_outlook_with_file_path(file_path)
        else:
            print("Invalid file path!")
    else:
        print("No file path provided!")

if __name__ == "__main__":
    main()
