from docxtpl import DocxTemplate
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
import os
import csv
import json

def create_mail_merge_docx():
    """Create a mail merge document using a DOCX template and data."""
    template_path = "template.docx"
    doc = DocxTemplate(template_path)
    if not os.path.exists(template_path):
        print("Error: Template file not found!")
        return

    data_file_path = "data.json"
    if not os.path.exists(data_file_path):
        print("Error: Data file not found!")
        return 
    elif data_file_path.endswith('.json'):
        try:
            context = []
            with open(data_file_path, 'r') as f:
                context = json.load(f)
            for entry in context:
                doc.render(entry)
                output_filename = "./output/{}_{}.docx".format(entry.get('id','output'), entry.get('name', 'output'))
                doc.save(output_filename)
                print(f"Mail merge document '{output_filename}' created successfully!")
        except Exception as e:
            print(f"Error creating mail merge document: {e}")

def send_bulk_emails_with_mail_merge():
    """Create mail merge documents and send them as email attachments with HTML template body."""
    try:
        sender_email = ""
        sender_password = "".replace("\xa0", "")
        subject = input("Enter the email subject: ")

        # Load HTML template for email body
        html_template_path = "template.html"
        if not os.path.exists(html_template_path):
            print("Error: HTML template file not found!")
            return

        with open(html_template_path, 'r', encoding='utf-8') as f:
            html_template = f.read()

        # Create the output directory if it doesn't exist
        if not os.path.exists("output"):
            os.makedirs("output")

        # Load data for mail merge
        data_file_path = "data.json"
        if not os.path.exists(data_file_path):
            print("Error: Data file not found!")
            return

        template_path = "template.docx"
        if not os.path.exists(template_path):
            print("Error: Template file not found!")
            return

        # Create mail merge documents and send emails
        try:
            with open(data_file_path, 'r') as f:
                context = json.load(f)

            count = 0
            for entry in context:
                # Create mail merge document
                doc = DocxTemplate(template_path)
                doc.render(entry)
                output_filename = "./output/{}_{}.docx".format(entry.get('id','output'), entry.get('name', 'output'))
                doc.save(output_filename)

                # Get recipient email (assuming 'email' field in data)
                receiver_email = entry.get('email') or entry.get('Email')
                if not receiver_email:
                    print(f"Warning: No email found for {entry.get('name', 'unknown')}. Skipping...")
                    continue

                # Personalize HTML body with data
                personalized_body = html_template
                for key, value in entry.items():
                    personalized_body = personalized_body.replace(f"{{{{{key}}}}}", str(value))

                # Create email message with HTML content
                msg = MIMEMultipart('alternative')
                msg['From'] = sender_email
                msg['To'] = receiver_email
                msg['Subject'] = Header(subject, 'utf-8')
                msg.attach(MIMEText(personalized_body, 'html', 'utf-8'))

                # Attach the generated mail merge document
                try:
                    part = MIMEBase('application', 'octet-stream')
                    with open(output_filename, 'rb') as attachment:
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(output_filename))
                    msg.attach(part)
                except Exception as e:
                    print(f"Error attaching document for {receiver_email}: {e}")
                    continue

                # Send email
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(sender_email, sender_password.encode("utf-8").decode("utf-8"))
                    server.send_message(msg)
                    server.quit()
                    count += 1
                    print(f"Email with mail merge document sent to {receiver_email}!")
                except smtplib.SMTPAuthenticationError:
                    print("Error: Authentication failed. Check your email and password.")
                    break
                except Exception as e:
                    print(f"Error sending email to {receiver_email}: {e}")

            print(f"\nMail merge & email sending completed! {count} emails sent.")
        except Exception as e:
            print(f"Error: {e}")
    except Exception as e:
        print(f"Error: {e}")

def send_email():
    """Send an email with optional attachment."""
    try:
        sender_email = input("Enter your email address: ")
        sender_password = input("Enter your email password (or app-specific password): ")
        receiver_email = input("Enter the recipient's email address: ")
        subject = input("Enter the email subject: ")
        body = input("Enter the email body: ")
        attachment_path = input("Enter the path to the attachment (leave blank if no attachment): ")

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = Header(subject, 'utf-8')

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if attachment_path and os.path.exists(attachment_path):
            try:
                part = MIMEBase('application', 'octet-stream')
                with open(attachment_path, 'rb') as attachment:
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                msg.attach(part)
            except Exception as e:
                print(f"Error attaching file: {e}")
                return

        # Send the email
        try:
            # Use Gmail SMTP server (you can change this for other email providers)
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            server.quit()
            print(f"Email sent successfully to {receiver_email}!")
        except smtplib.SMTPAuthenticationError:
            print("Error: Authentication failed. Check your email and password.")
        except smtplib.SMTPException as e:
            print(f"Error: Failed to send email: {e}")
        except Exception as e:
            print(f"Error: {e}")
    except Exception as e:
        print(f"Error: {e}")

def send_bulk_emails():
    """Send emails with merged documents to multiple recipients."""
    try:
        sender_email = input("Enter your email address: ")
        sender_password = input("Enter your email password (or app-specific password): ")
        
        # Load recipient data from CSV
        data_file = input("Enter the path to your recipient data file (CSV): ")
        if not os.path.exists(data_file):
            print("Error: Data file not found!")
            return
        
        subject = input("Enter the email subject: ")
        body = input("Enter the email body: ")
        attachment_path = input("Enter the path to the attachment file: ")
        
        if not os.path.exists(attachment_path):
            print("Error: Attachment file not found!")
            return
        
        # Read CSV and send emails
        with open(data_file, 'r') as f:
            reader = csv.DictReader(f)
            count = 0
            for row in reader:
                receiver_email = row.get('email') or row.get('Email')
                if not receiver_email:
                    print("Warning: No email field found in CSV. Using first column.")
                    receiver_email = list(row.values())[0]
                
                # Personalize body with data from CSV
                personalized_body = body
                for key, value in row.items():
                    personalized_body = personalized_body.replace(f"{{{{{key}}}}}", value)
                
                msg = MIMEMultipart()
                msg['From'] = sender_email
                msg['To'] = receiver_email
                msg['Subject'] = Header(subject, 'utf-8')
                msg.attach(MIMEText(personalized_body, 'plain', 'utf-8'))
                
                # Attach file
                try:
                    part = MIMEBase('application', 'octet-stream')
                    with open(attachment_path, 'rb') as attachment:
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
                    msg.attach(part)
                except Exception as e:
                    print(f"Error attaching file: {e}")
                    continue
                
                # Send email
                try:
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(sender_email, sender_password)
                    server.send_message(msg)
                    server.quit()
                    count += 1
                    print(f"Email sent to {receiver_email}")
                except Exception as e:
                    print(f"Error sending email to {receiver_email}: {e}")
        
        print(f"\nBulk email sending completed! {count} emails sent.")
    except Exception as e:
        print(f"Error: {e}")

def display_menu():
    """Display the main options menu."""
    print("\n" + "="*50)
    print("       MAIL MERGE & EMAIL APPLICATION")
    print("="*50)
    print("1. Create Mail Merge Document")
    print("2. Send Single Email")
    print("3. Send Bulk Emails")
    print("4. Send Mail Merge Documents via Email")
    print("5. Exit")
    print("="*50)

def main():
    """Main application loop with menu."""
    while True:
        display_menu()
        choice = input("Select an option (1-5): ")
        
        if choice == '1':
            print("\n--- Create Mail Merge Document ---")
            create_mail_merge_docx()
        elif choice == '2':
            print("\n--- Send Single Email ---")
            send_email()
        elif choice == '3':
            print("\n--- Send Bulk Emails ---")
            send_bulk_emails()
        elif choice == '4':
            print("\n--- Send Mail Merge Documents via Email ---")
            send_bulk_emails_with_mail_merge()
        elif choice == '5':
            print("Thank you for using Mail Merge & Email Application. Goodbye!")
            break
        else:
            print("Invalid option. Please select 1-5.")

if __name__ == "__main__":
    main()
