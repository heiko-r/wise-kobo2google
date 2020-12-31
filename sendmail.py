# Python script to send emails.

import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from email.utils import make_msgid

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False

ENCRYPTION_STARTTLS = 0
ENCRYPTION_TLS = 1

'''
Function to connect to SMTP server
'''
def connect_smtp(server, port, encryption, user, password):
    # Validate value for encryption
    if encryption != ENCRYPTION_STARTTLS and encryption != ENCRYPTION_TLS:
        print(f"Error in connect_smtp(): Invalid value for 'encryption': {encryption}.")
        return None
    
    # Connect to SMTP server
    debug_msg("Connecting to SMTP server")
    context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    try:
        if encryption == ENCRYPTION_STARTTLS:
            conn = smtplib.SMTP(server, port)
            conn.starttls(context=context)
        else:
            conn = smtplib.SMTP_SSL(server, port, context=context)
    except smtplib.SMTPException as error:
        print(f"Error: Exception caught while connecting to SMTP server, error: '{str(error)}'")
        return None
    
    # Authenticate
    debug_msg("Authenticating to SMTP server")
    try:
        conn.login(user, password)
    except smtplib.SMTPAuthenticationError as error:
        printf(f"Error: Invalid credentials for SMTP server, error: {str(error)}")
        return None
    except smtplib.SMTPException as error:
        printf(f"Error: Exception caught while logging into SMTP server, error: {str(error)}")
        return None
    
    return conn

'''
Function to disconnect from SMTP server
'''
def disconnect_smtp(conn):
    if conn:
        debug_msg("Closing SMTP connection")
        return conn.quit()
    else:
        return None

'''
Function to send a multipart plain text and HTML email
'''
def send_email(conn, fromAddr, toAddr, subject, txt, html):
    
    # Prepare multipart plain text/html message
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = fromAddr
    message["To"] = toAddr
    message["Date"] = formatdate(localtime = True)
    message["Message-ID"] = make_msgid(domain=fromAddr.split('@')[1])
    
    part1 = MIMEText(txt, "plain")
    part2 = MIMEText(html, "html")
    message.attach(part1)
    message.attach(part2)

    # Send the message
    debug_msg(f"Sending email to {toAddr}")

    try:
        recipients =  conn.sendmail(
            fromAddr, toAddr, message.as_string()
        )
        return recipients
    except smtplib.SMTPException as error:
        print(f"Error: Exception caught while sending email, error '{str(error)}'")
        return None
    
    # Failed
    return None


'''
Prints verbose debug message.
'''
def debug_msg(msg):
    if g_EnableDebugMsg:
        print(msg)