import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header


def SendToKindle(mail_host, mail_user, mail_pass, receiver, fullpath, bookname):
    try:
        message = MIMEMultipart()
        message['From'] = Header("SendToKindle", 'utf-8')
        message['To'] = receiver
        message['Subject'] = Header('convert')

        msg_attachment = MIMEApplication(open(fullpath, 'rb').read())
        msg_attachment.add_header('Content-Disposition', 'attachment', filename=bookname)
        message.attach(msg_attachment)

        server = smtplib.SMTP(mail_host, 587)
        server.ehlo()
        server.starttls()
        server.login(mail_user, mail_pass)
        server.sendmail(mail_user, [receiver], message.as_string())
        server.quit()
        return True
    except Exception as e:
        print("There is a exception", e)
        return False
