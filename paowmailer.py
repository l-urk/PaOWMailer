import os
import sys
import time
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
try:
    while True:
        def clear_screen():
            os.system('cls' if os.name == 'nt' else 'clear')
        clear_screen()
        banner = """\033[97m
         
        ██████╗  █████╗  ██████╗ ██╗    ██╗███╗   ███╗ █████╗ ██╗██╗     ███████╗██████╗  
        ██╔══██╗██╔══██╗██╔═══██╗██║    ██║████╗ ████║██╔══██╗██║██║     ██╔════╝██╔══██╗ 
        ██████╔╝███████║██║   ██║██║ █╗ ██║██╔████╔██║███████║██║██║     █████╗  ██████╔╝ 
        ██╔═══╝ ██╔══██║██║   ██║██║███╗██║██║╚██╔╝██║██╔══██║██║██║     ██╔══╝  ██╔══██╗   
        ██║     ██║  ██║╚██████╔╝╚███╔███╔╝██║ ╚═╝ ██║██║  ██║██║███████╗███████╗██║  ██║ 
        ╚═╝     ╚═╝  ╚═╝ ╚═════╝  ╚══╝╚══╝ ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚══════╝╚══════╝╚═╝  ╚═╝ 
        -- Python and Outlook Web Mailer -- created by l_, gitub.com/l-urk --                              
        """
        print(banner)
        def loading_spinner():
            symbols = ['\033[94m','🖥️✉🌐','🖥 ✉🌐','🖥  ✉🌐','🖥   ✉🌐','🖥    ✉🌐']
            for _ in range(10):
                for symbol in symbols:
                    print(f'\r{symbol}', end='', flush=True)
                    time.sleep(0.1)
            clear_screen()
            print(banner)
        clear_screen()
        print(banner)
        def display_menu():  
            print("\033[97m1(space). Send emails with pasted email list and message")
            print("2(enter). Send emails with contacts.csv and message.txt")
            print("3(ctrlC). Exit")        
            choice = input(":")
            if choice == "1" or choice == " ":
                clear_screen()
                print(banner)
                login_send_with_input()
            if choice == "2" or choice == "":
                clear_screen()
                print(banner)
                login_send_with_files()
            if choice == "3":
                clear_screen()
                sys.exit()
            else:
                clear_screen()
                print(banner)
                display_menu() 
            return choice
        def login_send_with_input():
            print("\033[92mLogin to your Outlook account.")
            sender_email = input("\033[97mEnter your Outlook email: ")
            sender_password = input("Enter your Outlook password: ")
            print("\033[94mAttempting to login...")
            loading_spinner()
            if sender_email == "":
                clear_screen()
                print(banner)
                display_menu()
            smtp_server = 'smtp-mail.outlook.com'
            smtp_port = 587
            smtp = smtplib.SMTP(smtp_server, port=smtp_port)
            smtp.starttls()
            while True:
                try:
                    smtp.login(sender_email, sender_password)
                    break
                except smtplib.SMTPException:
                    clear_screen()
                    print(banner)
                    print("\033[91mIncorrect username or password.")
                    time.sleep(1)
                    display_menu()   
            recipient_details = []
            while True:
                print(f'\033[32mLogged in as {sender_email}')
                print("\033[37mPlease add recipeints to your batch email.")
                print("May include recipient details in format \033[1m'First Last, Email First Last, Email'\033[0m")
                print("Or, may include recipient details as... \033[1m'Name, Email Name, Email...'\033[0m")
                print("For example: '\033[97mFoo Bar, \033[34mfoobar@foobar.com\033[97m, Rue L Ryuzaki, \033[34mb.b.rue.l.ryuzaki@gmail.com\033[97m... etc, more...")
                recipient_details = []
                    
                while True:
                    recipient_input = input("Enter a recipient as 'Name, Email' (or type 'done' to finish): ")

                    if recipient_input.lower() == 'done':
                        break
                    elif recipient_input.lower() == '':
                        clear_screen()
                        print(banner)
                        display_menu()

                    try:
                        recipient_name, recipient_email = recipient_input.split(',')
                        recipient_details.append((recipient_name.strip(), recipient_email.strip()))
                    
                        
                    except ValueError:
                        clear_screen()
                        print(banner)
                        print("\033[91mInvalid input. Please use 'Name, Email' format.")
                        display_menu()


                subject = input("Enter the subject: ")
                print("Enter the message text.")
                print("May include refernces like {recipient_name}, {recipient_email}, {sender_email}")
                print("For example: Hello, {recipient_name} of the , I trust you are aware of the significance of our mission.")
                message = input("In unity and obscurity, In unity and obscurity, Mr. Robot {sender_email}: ")
                break
            for name, recipient_email in recipient_details:
                try:
                    full_message = f"Subject: {subject}\n\nDear {name},\n{message}"
                    print(f'\033[32mEmail sent to {recipient_name} ({recipient_email}) successfully.')
                except Exception as e:
                    full_message = f"Subject: {subject}\n\nDear {name},\n{message}"
                    smtp.sendmail(sender_email, recipient_email, full_message)
                time.sleep(5)        
            smtp.quit()
        def login_send_with_files():
            sender_email = input("Enter your Outlook email: ")
            sender_password = input("Enter your Outlook password: ")
            smtp_server = 'smtp-mail.outlook.com'
            smtp_port = 587
            
            if not os.path.isfile('message.txt'):
                print("\033[31mThe 'message.txt' file does not exist. Please create the message.txt file and try again.")
                time.sleep(3)
                return # Exit the function

            with open('message.txt', 'r') as file:
                lines = file.read().split('\n')
                subject = lines[0]
                message = '\n'.join(lines[1:])
                recipient_details = []
                
            if not os.path.isfile('contacts.csv'):
                print("\033[31mThe 'contacts.csv' file does not exist. Please create the contacts.csv file and try again.")
                time.sleep(3)
                return # Exit the function    
                
            with open('contacts.csv', 'r') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    recipient_details.append(row)
            try:
                smtp = smtplib.SMTP(smtp_server, port=smtp_port)
                smtp.starttls()
                smtp.login(sender_email, sender_password)
                for recipient in recipient_details:
                    recipient_name = recipient['Name']
                    recipient_email = recipient['Email']
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient_email
                    msg['Subject'] = subject.replace('{recipient_name}', recipient_name)
                    msg.attach(MIMEText(message.replace('{recipient_name}', recipient_name), 'plain'))
                    try:
                        smtp.sendmail(sender_email, recipient_email, msg.as_string())
                        print(f'\033[32mEmail sent to {recipient_name} ({recipient_email}) successfully.')
                    except Exception as e:
                        print(f'Email to {recipient_name} ({recipient_email}) failed. Error: {str(e)}')
            except smtplib.SMTPException as e:
                clear_screen()
                print(banner)
                print("\033[91mIncorrect username or password.")
                time.sleep(1)
                display_menu()
            time.sleep(5)
        display_menu()
except KeyboardInterrupt:
    print("\033[97m")
    clear_screen()
