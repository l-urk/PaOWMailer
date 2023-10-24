import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import csv

# Email configuration
sender_email = input("Enter your Outlook email: ")
sender_password = input("Enter your Outlook password: ")
smtp_server = 'smtp-mail.outlook.com'
smtp_port = 587

# Read the subject and message from a text file
with open('message.txt', 'r') as file:
    lines = file.read().split('\n')

subject = lines[0]  # The first line is the subject
message = '\n'.join(lines[1:])  # The rest of the lines are the message

# Create a list to store the recipient details
recipient_details = []

# Read recipient details from the CSV file
with open('contacts.csv', 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        recipient_details.append(row)

# Connect to the Outlook SMTP server
smtp = smtplib.SMTP(smtp_server, port=smtp_port)
smtp.starttls()
smtp.login(sender_email, sender_password)

for recipient in recipient_details:
    recipient_name = recipient['Name']
    recipient_email = recipient['Email']

    # Create an email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject.replace('{recipient_name}', recipient_name)
    msg.attach(MIMEText(message.replace('{recipient_name}', recipient_name), 'plain'))

    try:
        # Send the email
        smtp.sendmail(sender_email, recipient_email, msg.as_string())
        print(f'Email sent to {recipient_name} ({recipient_email}) successfully.')
    except Exception as e:
        print(f'Email to {recipient_name} ({recipient_email}) failed. Error: {str(e)}')

# Quit the server
smtp.quit()

print('All emails sent successfully.')
