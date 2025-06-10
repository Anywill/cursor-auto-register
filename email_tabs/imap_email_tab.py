from email_tabs.email_tab_interface import EmailTabInterface
import imaplib
import email
import re
from email.header import decode_header
import time

class IMAPEmailTab(EmailTabInterface):
    def __init__(self):
        self.host = "imap.163.com"
        self.email_address = 'mcdonald2025@163.com'
        self.auth_code = 'RCSASnZggNgrjPha'
        self.imap = self.login()
        self.latest_uid = None

    def login(self):
        """连接网易邮箱服务器并登录"""
        try:
            # 使用 SSL 连接
            server = imaplib.IMAP4_SSL(self.host)
            server.login(self.email_address, self.auth_code)

            # 解决网易邮箱报错
            imaplib.Commands["ID"] = ('AUTH',)
            args = ("name", self.email_address, "contact", self.email_address, "version", "1.0.0", "vendor", "myclient")
            server._simple_command("ID", str(args).replace(",", "").replace("\'", "\""))

            return server
        except Exception as e:
            print(f"登录邮箱时出错: {str(e)}")
            raise

    def refresh_inbox(self):
        """刷新收件箱"""
        try:
            status, _ = self.imap.select('INBOX')
            if status != 'OK':
                print("无法选中 INBOX 文件夹")
                return False
            return True
        except Exception as e:
            print(f'选中 INBOX 文件夹时出错: {str(e)}')
            return False

    def check_for_cursor_email(self):
        """检查是否有新的 Cursor 邮件"""
        max_attempts = 5
        retry_interval = 10  # 10秒间隔

        for attempt in range(max_attempts):
            print(f"[调试] 第 {attempt + 1} 次尝试检查邮件...")

            if not self.refresh_inbox():
                print("[调试] 刷新收件箱失败")
                if attempt < max_attempts - 1:
                    print(f"[调试] 等待 {retry_interval} 秒后重试...")
                    time.sleep(retry_interval)
                continue

            try:
                status, messages = self.imap.search(None, 'UNSEEN')
                print(f"[调试] 搜索未读邮件状态: {status}, 消息: {messages}")
                if status != 'OK':
                    print("[调试] 搜索邮件失败")
                    if attempt < max_attempts - 1:
                        print(f"[调试] 等待 {retry_interval} 秒后重试...")
                        time.sleep(retry_interval)
                    continue

                for num in messages[0].split():
                    status, msg_data = self.imap.fetch(num, '(RFC822)')
                    print(f"[调试] fetch状态: {status}, 邮件编号: {num}")
                    if status != 'OK':
                        continue

                    msg = email.message_from_bytes(msg_data[0][1])
                    subject = self._decode_header(msg['Subject'])
                    from_header = self._decode_header(msg['From'])
                    print(f"[调试] 邮件主题: {subject}, 发件人: {from_header}")

                    if "Cursor" in subject or "Cursor" in from_header:
                        self.latest_uid = num
                        print(f"[调试] 找到Cursor相关邮件，编号: {num}")
                        return True

                    if "gmail" in subject or "gmail" in from_header:
                        self.latest_uid = num
                        print(f"[调试] 找到gmail相关邮件，编号: {num}")
                        return True

                print("[调试] 未找到包含Cursor的未读邮件")
                if attempt < max_attempts - 1:
                    print(f"[调试] 等待 {retry_interval} 秒后重试...")
                    time.sleep(retry_interval)
                continue

            except Exception as e:
                print(f"[调试] 检查邮件时出错: {str(e)}")
                if attempt < max_attempts - 1:
                    print(f"[调试] 等待 {retry_interval} 秒后重试...")
                    time.sleep(retry_interval)
                continue

        print("[调试] 达到最大尝试次数，未找到Cursor相关邮件")
        return False

    def get_verification_code(self):
        """获取验证码"""
        if not self.latest_uid:
            print("[调试] 没有最新邮件编号")
            return ""

        try:
            status, msg_data = self.imap.fetch(self.latest_uid, '(RFC822)')
            print(f"[调试] fetch验证码邮件状态: {status}")
            if status != 'OK':
                return ""

            msg = email.message_from_bytes(msg_data[0][1])
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        if part.get_charsets() is None:
                            body = part.get_payload(decode=True)
                        else:
                            body = part.get_payload(decode=True).decode(part.get_charsets()[0])
                        if isinstance(body, bytes):
                            body = body.decode('utf-8', errors='ignore')
                        print(f"[调试] 邮件正文内容: {body}")  # 打印前200字符
                        patterns = [
                            r"\b\d{6}\b",  # 标准6位数字
                            r"code[:\s]+(\d{6})",  # 包含"code"的6位数字
                            r"verification[:\s]+(\d{6})",  # 包含"verification"的6位数字
                            r"(\d{6})[^\d]*$"  # 行尾的6位数字
                        ]

                        for pattern in patterns:
                            match = re.search(pattern, body, re.IGNORECASE)
                            if match:
                                print(match)
                                code = match.group(1) if len(match.groups()) > 0 else match.group(0)
                                print(f"[调试] 匹配到验证码: {code}")
                                return code
                    except Exception as e:
                        print(f"[调试] 解析邮件内容时出错: {str(e)}")
                        continue
            print("[调试] 未在邮件正文中找到验证码")
            return ""
        except Exception as e:
            print(f"[调试] 获取验证码时出错: {str(e)}")
            return ""

    def _decode_header(self, header):
        """解码邮件头信息"""
        try:
            decoded_header = decode_header(header)[0]
            if isinstance(decoded_header[0], bytes):
                return decoded_header[0].decode(decoded_header[1] or 'utf-8', errors='ignore')
            return decoded_header[0]
        except Exception as e:
            print(f"解码邮件头信息时出错: {str(e)}")
            return ""

    def __del__(self):
        """析构函数，确保连接正确关闭"""
        try:
            if hasattr(self, 'imap'):
                self.imap.close()
                self.imap.logout()
        except:
            pass
