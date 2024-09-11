import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def email_send_function(to_, subj_, msg_, from_, pass_):
    try:
        # Use Zoho's SMTP server
        s = smtplib.SMTP("smtp.zoho.com", 587)
        s.starttls()
        s.login(from_, pass_)

        # Create a MIMEMultipart email message
        email_message = MIMEMultipart()
        email_message['From'] = from_
        email_message['To'] = to_
        email_message['Subject'] = subj_

        # Attach the email body with proper encoding (UTF-8)
        body = MIMEText(msg_, 'plain', 'utf-8')
        email_message.attach(body)

        # Convert the message to a string and send it
        s.sendmail(from_, to_, email_message.as_string())

        # Close the SMTP connection
        s.quit()

        return "s"  # Success

    except smtplib.SMTPAuthenticationError:
        print("Authentication Error: Unable to log in. Please check your email and password.")
        return "f"  # Failure

    except Exception as e:
        print(f"Failed to send email to {to_}. Error: {e}")
        return "f"  # Failure
