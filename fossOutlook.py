import os

import win32com.client


# Function to log in to Microsoft Outlook
def login(username, password):
    # Create a Microsoft Outlook object
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Log in to your Microsoft Outlook account
    outlook.Session.Logon(username, password)

    return outlook


# Function to read emails from a specified folder
def read_emails(outlook, folder_name):
    # Get the specified folder from Microsoft Outlook
    # folder = outlook.Session.Folders[folder_name]
    # folder=outlook.Session.GetDefaultFolder(6).Folders[1].Folders[1]
    folder = outlook.Session.GetDefaultFolder(6).Folders["Impotent"].Folders["Amit Datar"]

    # Iterate through the emails in the folder
    for email in folder.Items:
        # Return the email object
        yield email


# Function to save an email and its attachments to a specified directory
def save_email(email, save_dir):
    # Save the email to the specified directory
    # email.SaveAs(f'"{save_dir}{email.Subject}.msg"')
    save_email_path=f'"{save_dir}{email.Subject}.msg"'
    with open(att_name, 'wb') as fl:
        fl.write(att.data)
    # Iterate through the attachments in the email
    for attachment in email.Attachments:
        # Save the attachment to the specified directory
        attachment.SaveAsFile(save_dir + "/" + attachment.Filename)


# Main program
if __name__ == "__main__":
    # Log in to Microsoft Outlook
    username = "*****"
    password = "****"
    outlook = login(username, password)

    # Read emails from the Inbox folder
    folder_name = "Inbox"
    directory_path = os.getcwd()
    SOURCE_FOLDER = directory_path + "\\sources\\"

    save_dir = SOURCE_FOLDER
    for email in read_emails(outlook, folder_name):
        save_email(email, save_dir)

    # Close the Microsoft Outlook object
    outlook.Quit()

