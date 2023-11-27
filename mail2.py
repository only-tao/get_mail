import os
import imaplib
import email
from email.header import decode_header
import tkinter as tk
from tkinter import ttk
def fetch_and_save_emails(work_order,save_path):
    if not os.path.exists(save_path):
        print(f"Path {save_path} doesn't exist.")
        os.makedirs(save_path)

    cn_num = {0: '零', 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}
    character_order = cn_num.get(int(work_order), str(work_order))

    num = 0
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

        # get subject and encoding 
        subject, encoding = decode_header(msg["Subject"])[0]
        # 解码邮件主题
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")
        print(f"Subject: {subject}")

        if f"第{work_order}次" in subject or f"第{character_order}次" in subject:
            num += 1
            print("#########")
            print(f"select_Subject: {subject}")
            # 检查邮件是否包含附件
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get("Content-Disposition") is not None:
                        filename = part.get_filename()
                        if filename:
                            # 保存附件到本地文件
                            filepath = os.path.join(save_path, f"{subject}.pdf")
                            print(f"filepath: {filepath}")
                            print("#########")
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))

    print(f"\nYou get {num} mails in work{work_order}")

    # 关闭连接
    mail.logout()

# Example usage:
# fetch_and_save_emails(work_order,save_path)
# 定义一个函数，用于处理按钮点击事件
work_order = "4"
save_path = f"/home/tao/Desktop/embedded_system/work{work_order}-stu/"
def on_fetch_button_click():
    work_order = work_order_entry.get()
    save_path = save_path_entry.get()
    fetch_and_save_emails(work_order, save_path)

# 创建主窗口
root = tk.Tk()
root.title("Email Fetcher")
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)
entry_width = max(len(save_path),len(work_order))+50
# Set a larger font size
font_size = 14
font = ("Helvetica", font_size)
# 在窗口中添加标签、输入框和按钮
work_order_label = ttk.Label(root, text="工单号:", font=font)
work_order_label.grid(row=0, column=0, padx=5, pady=5, sticky="E")

work_order_entry = ttk.Entry(root,width=entry_width, font=font)
work_order_entry.grid(row=0, column=1, padx=5, pady=5, sticky="W")
work_order_entry.insert(0,work_order)  # 设置默认值

save_path_label = ttk.Label(root, text="保存路径:", font=font)
save_path_label.grid(row=1, column=0, padx=5, pady=5, sticky="E")

save_path_entry = ttk.Entry(root,width=entry_width, font=font)
save_path_entry.grid(row=1, column=1, padx=5, pady=5, sticky="W")
save_path_entry.insert(0,save_path)
fetch_button = ttk.Button(root, text="Fetch Emails", command=on_fetch_button_click)
fetch_button.grid(row=2, column=0, columnspan=2, pady=10)

# Center the window on the screen
root.update_idletasks()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = root.winfo_reqwidth()
window_height = root.winfo_reqheight()
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
root.geometry(f"+{x_position}+{y_position}")

# 启动主循环
root.mainloop()