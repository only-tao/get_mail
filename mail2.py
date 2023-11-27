import imaplib
import email
from email.header import decode_header
import os
import re
work_order = "4"
save_path = "/home/tao/Desktop/embedded_system/work3-stu/"
cn_num = {0: '零', 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}
character_order = cn_num[int(work_order)]
# print(character_order)
num =0
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

# 选择邮箱中的收件箱git
mail.select("inbox")

# 搜索邮件标题包含 "实验" 的邮件

status, email_ids = mail.search(None, "ALL")
email_ids = email_ids[0].split()

for email_id in email_ids:
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    #get subject and encoding 
    subject, encoding = decode_header(msg["Subject"])[0]
      # 解码邮件主题
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8")
    print(f"all_Subject: {subject}")
    
    if f"第{work_order}次" in subject or f"第{character_order}次" in subject:
        num +=1 
        print("#########")
        print(f"select_Subject: {subject}")
        # 检查邮件是否包含附件
        if msg.is_multipart():
            for part in msg.walk():
                if part.get("Content-Disposition") is not None:
                    filename = part.get_filename()
                    if filename:
                        # 保存附件到本地文件
                        filepath = os.path.join(save_path, subject + ".pdf")
                        print("filepath", filepath)
                        print("#########")
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
print(f"You get {num} mails in work{work_order}")
# 关闭连接
mail.logout()
