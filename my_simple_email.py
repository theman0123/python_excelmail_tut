import smtplib

from email.mime.text import MIMEText

gmail_user = "yourUserName@gmail.com"
gmail_appPassword = "yourAppPassword"

sent_from = ['yourUserName@gmail.com']
to = ['probablyYourUserNameForNow@gmail.com']
text = 'you owe me a million dollars, bro'


msg = MIMEText(text)

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login(gmail_user, gmail_appPassword)
server.sendmail(sent_from, to, msg.as_string())
server.quit()

#tested-works#
