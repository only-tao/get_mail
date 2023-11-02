import imaplib
import email
from email.header import decode_header
import os
# 设置你的电子邮件凭证
with open("passwd.txt", "r") as file:
    lines = file.readlines()

    if len(lines) >= 2:
        email_user = lines[0].strip()  # 去除行尾的换行符和空格
        email_pass = lines[1].strip()  # 去除行尾的换行符和空格
    else:
        print("文件test.txt中的行数不足2行")
        exit()

# 连接到腾讯企业邮 IMAP 服务器
mail = imaplib.IMAP4_SSL("imap.exmail.qq.com")

# 登录到你的邮箱帐户
mail.login(email_user, email_pass)

# 选择邮箱中的收件箱
mail.select("inbox")

# 搜索邮件标题包含 "实验" 的邮件
# search_criteria = '(SUBJECT "=?UTF-8?B?' + '实验'.encode('utf-8').hex() + '?=")'

status, email_ids = mail.search(None, "ALL")
email_ids = email_ids[0].split()

for email_id in email_ids:
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)
    subject, encoding = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8")
    print(f"Subject: {subject}")
    # for part in msg.walk():
    #     if part.get_content_maintype() == "multipart":
    #         continue
    #     if part.get("Content-Disposition") is None:
    #         continue

    #     filename = part.get_filename()
    #     if filename is not None:
    #         filename = decode_header(filename)[0][0]
    #         if filename is not None:
    #             print(f"Downloading attachment: {filename}")
    #             file_data = part.get_payload(decode=True)
    #             # 将附件保存到本地
    #             with open(filename, "wb") as f:
    #                 f.write(file_data)

    # 检查邮件是否包含附件
    if msg.is_multipart():
        for part in msg.walk():
            if part.get("Content-Disposition") is not None:
                filename = part.get_filename()
                if filename:
                    # 保存附件到本地文件
                    # print("filename =",filename)
                    filepath = os.path.join("/data/Code/file_test", subject+".pdf")
                    print("filepath",filepath)
                    print("#########")
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))

# 关闭连接
mail.logout()
