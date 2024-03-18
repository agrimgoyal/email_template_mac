import pandas as pd
import subprocess

def send_email_apple_mail(email, subject, content, attachment_path):
    attachment_command = f'''
    tell application "Finder"
        set theAttachment to POSIX file "{attachment_path}"
    end tell
    ''' if attachment_path else ''

    applescript_command = f'''
    tell application "Mail"
        set newMessage to make new outgoing message with properties {{subject: "{subject}", content: "{content}", visible: true}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{address: "{email}"}}
            {attachment_command}
            if "{attachment_path}" is not "" then
                make new attachment with properties {{file name:theAttachment as alias}} at after the last paragraph
            end if
        end tell
        send newMessage
    end tell
    '''
    subprocess.run(["osascript", "-e", applescript_command])

# Path to the Excel file
file_path = 'marketing_sheet.xlsx'

# Read the Excel file
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Loop through each row to send emails
for index, row in df.iterrows():
    email = row['Mail']
    subject = row['Subject']
    content = row['Content']
    attachment = row['Attachment'] if 'Attachment' in row and not pd.isna(row['Attachment']) else ''

    send_email_apple_mail(email, subject, content, attachment)

print("Emails processing complete.")
