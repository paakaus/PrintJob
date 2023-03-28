import os
import shutil
import datetime
import win32api
import argparse

# Set the path of the directory containing the files
dir_path = '##%%Directory%%##'

# Define the command line arguments
parser = argparse.ArgumentParser(description='Print the run sheets.')
parser.add_argument('--test', action='store_true', help='run the script in test mode')
args = parser.parse_args()

# Get the current day of the week
day_of_week = datetime.datetime.today().weekday()

# Check if it's a weekday
if day_of_week in [0, 1, 2, 3]:
    # Print one of each Excel sheet and one of each Word document in Daily folder
    daily_folder = os.path.join(dir_path, "Daily")
    for file_name in os.listdir(daily_folder):
        if file_name.endswith(".xls") or file_name.endswith(".xlsx") or file_name.endswith(".doc") or file_name.endswith(".docx"):
            file_path = os.path.join(daily_folder, file_name)
            print("Printing", file_path)  # Print the file path instead of actually printing
            if not args.test:
                win32api.ShellExecute(0, "print", file_path, None, ".", 0)

elif day_of_week == 4:
    # Print three of each Excel sheet and one Word document in Daily folder
    daily_folder = os.path.join(dir_path, "Daily")
    for file_name in os.listdir(daily_folder):
        if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            file_path = os.path.join(daily_folder, file_name)
            for i in range(daily_copies):
                print("Printing", file_path)  # Print the file path instead of actually printing
        elif file_name.endswith(".doc") or file_name.endswith(".docx"):
            file_path = os.path.join(daily_folder, file_name)
            # Print one copy of the modified Word document in smaller font size
            # Open the Word document again and modify the font size
            with open(file_path, "r+") as file:
                text = file.read()
                text = text.replace("<w:sz w:val=\"24\"/>", "<w:sz w:val=\"12\"/>")
                file.seek(0)
                file.write(text)
                file.truncate()
            for i in range(daily_copies):
                print("Printing", file_path)  # Print the file path instead of actually printing
            
    # Print one of each Excel sheet and three of the Word document in Weekend folder
    weekend_folder = os.path.join(dir_path, "Weekend")
    for file_name in os.listdir(weekend_folder):
        if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            file_path = os.path.join(weekend_folder, file_name)
            for i in range(weekend_copies):
                print("Printing", file_path)  # Print the file path instead of actually printing
        elif file_name.endswith(".doc") or file_name.endswith(".docx"):
            file_path = os.path.join(weekend_folder, file_name)
            # Print one copy of the modified Word document in smaller font size
            # Open the Word document again and modify the font size
            with open(file_path, "r+") as file:
                text = file.read()
                text = text.replace("<w:sz w:val=\"24\"/>", "<w:sz w:val=\"12\"/>")
                file.seek(0)
                file.write(text)
                file.truncate()
            for i in range(weekend_copies):
                print("Printing", file_path)  # Print the file path instead of actually printing


# Ask the user if they need a refill on attachments
attachment_folder = os.path.join(dir_path, "Attachments")
refill = input("Do you need a refill on attachments? (yes or no): ").lower()
if refill == "yes":
    # Ask the user which attachment they want to print
    attachment_options = {
        1: "Incident_Report_Master.xls",
        2: "ParraPark_Combined_Master.xls",
        3: "Patrol_break.xlsx"
    }
    attachment_choice = input("Which attachment do you want to print? (1 Incident Reports, 2 Parramatta Park, or 3 Patrol Break): ")
    try:
        attachment_choice = int(attachment_choice)
        if attachment_choice in attachment_options:
            attachment_file = os.path.join(attachment_folder, attachment_options[attachment_choice])
            num_copies = int(input("How many copies do you want to print? "))
            for i in range(num_copies):
                print("Printing", attachment_file)
                if not is_test_mode:
                    win32api.ShellExecute(0, "print", attachment_file, None, ".", 0)  # Print the attachment
        else:
            print("Invalid choice.")
    except ValueError:
        print("Invalid input. Please enter a valid integer.")
print("Printing Starting")
             
