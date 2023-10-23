# PaOWMailer
Python and Outlook Web Mailer
PaOWMailer is a Python script that simplifies sending personalized bulk emails through your Outlook email account. It allows you to create customized messages and import recipient details from a CSV file, making it a convenient tool for efficiently reaching multiple contacts while maintaining a personal touch. Secure authentication and user-friendly design ensure a smooth email-sending experience for professionals and small businesses.
# Requirements:

# message.txt:
Description: This text file contains the email message that you want to send to the recipients. It includes both the subject and the body of the email, and it can contain placeholders for recipient-specific information.
Format: The file should be in plain text format and organized as follows:

Subject: Your Subject Line

Hello {recipient_name},
We are excited to announce our latest product launch...

# contacts.csv:

Description: This CSV (Comma-Separated Values) file contains the details of the recipients, including their names and email addresses. This file serves as the recipient list for your email campaign.
Format: The file should be structured with two columns: "Name" and "Email," where each row represents a recipient. The "Name" column should contain the recipient's name, and the "Email" column should contain their email address.

Name,Email
John Doe,johndoe@example.com
Jane Smith,janesmith@example.com
Ensure that you have these files in the specified formats and in the same location as locations as PaOWMailer.py as the program requires them for sending personalized bulk emails.
