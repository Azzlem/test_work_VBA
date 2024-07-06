import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path


def create_email(sender_email, recipient_email, subject, body, attachment_path):
    """Создание письма с вложением."""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    attachment = MIMEBase('application', 'octet-stream')
    with open(attachment_path, 'rb') as file:
        attachment.set_payload(file.read())

    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', f'attachment; filename={attachment_path.name}')
    msg.attach(attachment)

    return msg


def send_email(smtp_server, smtp_port, sender_email, password, recipient_email, msg):
    """Отправка письма через SMTP сервер."""
    try:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print("Email sent successfully")
    except smtplib.SMTPException as e:
        print(f"Error sending email: {e}")


def send():
    smtp_server = 'smtp.yandex.ru'
    smtp_port = 465
    sender_email = 'your_email@yandex.ru'
    password = 'your_password'
    recipient_email = 'recipient@example.com'
    subject = 'Список тем для доклада'
    body = 'Во вложении файл со списком тем и найденными источниками.'
    file_path = Path('TestTask2.xlsx')

    msg = create_email(sender_email, recipient_email, subject, body, file_path)
    send_email(smtp_server, smtp_port, sender_email, password, recipient_email, msg)



