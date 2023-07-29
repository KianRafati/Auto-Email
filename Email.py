import smtplib
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import socket
import time
import pandas as pd

# Function to send the personalized email with attachments


def send_email(sender_email, sender_password, recipient_name, recipient_email, file_path, png1_path, png2_path):
    subject = "Invitation to 'Synchronous 2023' Event"

    message = f"""Dear Prof. {recipient_name},
    
    I am writing on behalf of the Resana Association, the student organization affiliated with the Electrical Engineering Department of Sharif University of Technology; the highest-ranked university in Iran. Our aim is to link students more to industry and research by organizing cultural, scientific, and technological events.

    We are excited to announce our upcoming event, "Synchronous 2023", which will focus on trending projects and problems in Electrical Engineering. As a leading university in academia, we would like to invite you to participate and partner with us in this event by suggesting a challenge or project to our students.

    The students will solve the problem for you. We believe for this partnership to be mutually beneficial for both SUT students and your field of work and in honor of promoting research and students' scientific endeavor, by claiming rights to the solution our students provide to your problems, the members of the winning team will be provided with an internship position or a recommendation letter from you.

    If you are interested, we can discuss the details and terms of the event in a future interview. We would be honored to have your university's name associated with our event. Please find the attached proposal file for your reference.

    Thank you for considering our proposal. We look forward to hearing from you soon.

    Best regards,
    Kian Raf’ati Sajedi
    """

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(message, "plain"))

    # Attach the main file
    with open(file_path, 'rb') as file:
        file_data = file.read()
        file_name = file.name
        part = MIMEApplication(file_data)
        part.add_header('Content-Disposition', 'attachment', filename=file_name)
        msg.attach(part)

    # Attach the first PNG image
    with open(png1_path, 'rb') as file:
        file_data = file.read()
        part = MIMEImage(file_data)
        part.add_header('Content-Disposition', 'attachment', filename=png1_path)
        msg.attach(part)

    # Attach the second PNG image
    with open(png2_path, 'rb') as file:
        file_data = file.read()
        part = MIMEImage(file_data)
        part.add_header('Content-Disposition', 'attachment', filename=png2_path)
        msg.attach(part)

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print(
            f"Email sent to Prof. {recipient_name} ({recipient_email}) successfully.")
    except (socket.timeout, smtplib.SMTPException, OSError) as e:
        print("Error sending email:", e)
    except Exception as e:
        print(
            f"Error sending email to Prof. {recipient_name} ({recipient_email}): {e}")


# Your Outlook email credentials
sender_email = "name@example.com"
sender_password = "password"

# Read data from the Excel file (assuming it has 'Name' and 'Email' columns)
data = pd.read_excel('EE_ProfessorsDB.xlsx')

# Paths to files and images to be attached
file_path = "Synchron Proposal.pdf"
png1_path = "SUT-Logo.png"
png2_path = "Resana-Logo.png"

# Sending personalized emails with attachments to all recipients
for index, row in data.iterrows():
    name = row['Name']
    email = row['Email']
    send_email(sender_email, sender_password, name, email, file_path, png1_path, png2_path)
    time.sleep(5)  
