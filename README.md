![paowmailer](https://github.com/l-urk/PaOWMailer/assets/112792604/35c926af-409d-44dd-8104-0ff1a28050be)
# PaOWMailer
Python and Outlook Web Mailer
PaOWMailer is a Python script that simplifies sending personalized bulk emails through your Outlook email account. It allows you to create customized messages and import recipient details from a CSV file, making it a convenient tool for efficiently reaching multiple contacts while maintaining a personal touch. Secure authentication and user-friendly design ensure a smooth email-sending experience for professionals and small businesses.
# Requirements:

# message.txt:
This text file contains the email message that you want to send to the recipients. It includes both the subject and the body of the email, and it can contain placeholders for recipient-specific information. Using {recipient_name} will reference the name of the recipient being sent to.
The file should be in plain text format and organized as follows:

<img width="494" alt="image" src="https://github.com/l-urk/PaOWMailer/assets/112792604/36dbf696-260e-4c8d-8a6c-4162ac4240a6">

Be sure to change this to your own custom message before sending to any clients.

# contacts.csv:

Description: This CSV (Comma-Separated Values) file contains the details of the recipients, including their names and email addresses. This file serves as the recipient list for your email campaign.
Format: The file should be structured with two columns: "Name" and "Email," where each row represents a recipient. The "Name" column should contain the recipient's name, and the "Email" column should contain their email address.

<img width="205" alt="contacts" src="https://github.com/l-urk/PaOWMailer/assets/112792604/656e5c9e-3786-4204-bafe-914781067c6b">

Be sure to change these names and emails to your own list before attempting to send anything.

Ensure that you have these files in the specified formats and in the same location as locations as PaOWMailer.py as the program requires them for sending personalized bulk emails.
