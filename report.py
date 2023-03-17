import os
import shutil
import datetime
import win32api
import win32print

# Set the path of the directory containing the files
dir_path = 'C:\\Users\\PSS\\Desktop\\RunSheets\\'

# Get the current day of the week
day_of_week = datetime.datetime.today().weekday()

# Check if it's a weekday
if day_of_week in [0, 1, 2, 3]:
    # Print one of each Excel sheet and one of each Word document in Daily folder
    daily_folder = os.path.join(dir_path, "Daily")
    for file_name in os.listdir(daily_folder):
        if file_name.endswith(".xls") or file_name.endswith(".xlsx") or file_name.endswith(".doc") or file_name.endswith(".docx"):
            file_path = os.path.join(daily_folder, file_name)
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)

elif day_of_week == 4:
    # Print three of each Excel sheet and one Word document in Daily folder
    daily_folder = os.path.join(dir_path, "Daily")
    for file_name in os.listdir(daily_folder):
        if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            file_path = os.path.join(daily_folder, file_name)
            for i in range(3):
                win32api.ShellExecute(0, "print", file_path, None, ".", 0)
        elif file_name.endswith(".doc") or file_name.endswith(".docx"):
            file_path = os.path.join(daily_folder, file_name)
            # Open the Word document and replace the date with the current date
            with open(file_path, "r+") as file:
                text = file.read()
                today = datetime.date.today().strftime("%d/%m/%Y")
                text = text.replace("DATE", today)
                file.seek(0)
                file.write(text)
                file.truncate()
            # Print one copy of the modified Word document in smaller font size
            # Open the Word document again and modify the font size
            with open(file_path, "r+") as file:
                text = file.read()
                text = text.replace("<w:sz w:val=\"24\"/>", "<w:sz w:val=\"12\"/>")
                file.seek(0)
                file.write(text)
                file.truncate()
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)

    # Print one of each Excel sheet and three of the Word document in Weekend folder
    weekend_folder = os.path.join(dir_path, "Weekend")
    for file_name in os.listdir(weekend_folder):
        if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
            file_path = os.path.join(weekend_folder, file_name)
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
        elif file_name.endswith(".doc") or file_name.endswith(".docx"):
            file_path = os.path.join(weekend_folder, file_name)
            # Open the file and read the text
            with open(file_path, "r", encoding="iso-8859-1") as file:
                text = file.read()

            # Replace the date with the current date
            today = datetime.date.today().strftime("%d/%m/%Y")
            text = text.replace("DATE", today)

            # Write the modified text back to the file
            with open(file_path, "w", encoding="iso-8859-1") as file:
                file.write(text)
            # Print three copies of the modified Word document
            for i in range(3):
                win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            # Print one copy of the modified Word document in smaller font size
            # Open the Word document again and modify the font size
            with open(file_path, "r+", encoding="iso-8859-1") as file:
                text = file.read()
                text = text.replace("<w:sz w:val=\"24\"/>", "<w:sz w:val=\"12\"/>")
                file.seek(0)
                file.write(text)
                file.truncate()
            # Rename the file to include today's date
            today_str = datetime.date.today().strftime("%Y%m%d")
            new_file_name = f"{today_str}_{file_name}"
            new_file_path = os.path.join(weekend_folder, new_file_name)
            os.rename(file_path, new_file_path)
            # Print one copy of the modified Word document in smaller font size with the new file name
            win32api.ShellExecute(0, "print", new_file_path, None, ".", 0)


# Ask the user if they need a refill on attachments
attachment_folder = os.path.join(dir_path, "Attachments")
refill = input("Do you need a refill on attachments? (yes or no): ").lower()
if refill == "yes":
    # Ask the user which attachment they want to print
    attachment_options = {
        "1": "Incident_Report_Master.xls",
        "2": "ParraPark_Combined_Master.xls",
        "3": "Patrol_break.xls"
    }
    attachment_choice = input("Which attachment do you want to print? (1, 2, or 3): ")
    if attachment_choice in attachment_options:
        attachment_file = os.path.join(attachment_folder, attachment_options[attachment_choice])
        num_copies = int(input("How many copies do you want to print? "))
        for i in range(num_copies):
            win32api.ShellExecute(0, "print", attachment_file, None, ".", 0)
    else:
        print("Invalid choice.")
