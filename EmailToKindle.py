import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr


def smtp_text():
    """纯文本格式邮件"""
    from_address = "congcong.liu@hotmail.com"
    password = "Lcc.84562"

    # smtp 服务器地址
    smtp_address = "smtp.office365.com"
    # 目标地址
    to_address = "liucongcong1@cn.wilmar-intl.com"
    try:
        # 邮件内容
        msg = MIMEMultipart()

        msg['From'] = formataddr(["sender", from_address])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg['To'] = formataddr(["receiver", to_address])
        msg['Subject'] = 'Python SMTP 邮件测试' # 邮件的主题

        msg_text = MIMEText("This is a email test and by python.", "plain", "utf-8")
        msg.attach(msg_text)

        msg_attachment = MIMEApplication(open('', 'rb').read())
        msg_attachment.add_header('Content-Disposition', 'attachment', filename='怎样选择成长股.mobi')
        msg.attach(msg_attachment)

        server = smtplib.SMTP(smtp_address, 587)
        server.ehlo()  # 向Gamil发送SMTP 'ehlo' 命令
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, [to_address], msg.as_string())
        server.quit()
        print('邮件发送成功')
    except Exception as e:
        print("There is a exception", e)

if __name__ == '__main__':
    smtp_text()
