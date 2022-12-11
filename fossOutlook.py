
# Function to log in to Microsoft Outlook
def login(username, password):
    # Create a Microsoft Outlook object
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Log in to your Microsoft Outlook account
    outlook.Session.Logon(username, password)

    return outlook


# Function to read emails from a specified folder
def read_emails(outlook, folder_name):
    # Get the specified folder from Microsoft Outlook
    folder = outlook.Session.Folders[folder_name]

    # Iterate through the emails in the folder
    for email in folder.Items:
        # Return the email object
        yield email


# Function to save an email and its attachments to a specified directory
def save_email(email, save_dir):
    # Save the email to the specified directory
    email.SaveAs(save_dir + "/" + email.Subject + ".msg")

    # Iterate through the attachments in the email
    for attachment in email.Attachments:
        # Save the attachment to the specified directory
        attachment.SaveAsFile(save_dir + "/" + attachment.Filename)


# Main program
if __name__ == "__main__":
    # Log in to Microsoft Outlook
    username = "your_username"
    password = "your_password"
    outlook = login(username, password)

    # Read emails from the Inbox folder
    folder_name = "Inbox"
    for email in read_emails(outlook, folder_name):
        # Save the email and its attachments to the specified directory
        save_dir = "C:/email_backup"
        save_email(email, save_dir)

    # Close the Microsoft Outlook object
    outlook.Quit()

