import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(subject, body, to_email):
    # SMTP server configuration
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_user = 'webscrapertestnebelungen@gmail.com'
    smtp_password = 'orxa gayk uurh cdkk'

    # Create a MIMEText object
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    server = None  # Initialize the server variable

    try:
        # Establish a secure session with the server
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=60)  # Set a 60-second timeout
        server.starttls()  # Upgrade the connection to secure
        server.login(smtp_user, smtp_password)  # Login to the email account

        # Send the email
        server.sendmail(smtp_user, to_email, msg.as_string())
        print('Email sent successfully!')

    except smtplib.SMTPException as e:
        print(f'Failed to send email: {e}')

    except Exception as e:
        print(f'An error occurred: {e}')

    finally:
        if server is not None:
            server.quit()  # Ensure the server is quit only if it was initialized

# Usage
send_email('Test Subject', 'Test Body', 'steglichmaximilian@gmail.com')