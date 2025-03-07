from astrbot.api.event import filter, AstrMessageEvent, MessageEventResult
from astrbot.api.star import Context, Star, register
from astrbot.api import logger



@register("astrbot_plugin_mail", "mail", "一个邮件插件, 主要用于查询邮件", "1.0.2")
class MyPlugin(Star):
    def __init__(self, context: Context, config: dict):
        try:
            self.init_config(config)
        except Exception as e:
            logger.error(f"初始化配置失败: {e}")
            raise e

    def init_config(self, config: dict):
        """初始化配置"""
        self.mail = None
        self.config = config
        if(config["mail_address"] == "" or config["mail_password"] == ""):
            raise Exception("邮箱地址或密码未设置，请前往管理面板->插件管理->astrbot_plugin_mail 设置邮箱地址和密码")
        else:
            self.mail = self.login_mail()

    def login_mail(self):
        """登录邮箱"""
        if(self.mail is not None):
            return self.mail
        try:
            import imaplib
        except ImportError:
            raise Exception("imaplib 模块未安装，请前往管理面板->控制台->安装pip库 安装 imaplib 这个库")
        try:
            # 连接到 IMAP 服务器
            mail = imaplib.IMAP4_SSL(self.config["mail_host"], self.config["mail_port"])
            mail.login(self.config["mail_address"], self.config["mail_password"])  # 使用您的 QQ 邮箱和授权码

            return mail
        except Exception as e:
            logger.error(f"登陆失败: {e}")
            raise e
    
    def contains_keywords(self, msg, keywords):
        """检查邮件标题或正文中是否包含关键词"""
        # 获取标题
        subject = msg["Subject"] or ""
        
        # 获取正文（处理多种邮件格式）
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        # 检查关键词（不区分大小写）
        text_to_check = (subject + " " + body).lower()
        return any(keyword.lower() in text_to_check for keyword in keywords)

    def get_mail_attachments(self, msg):
        """获取指定邮件的附件"""
        try:
            # 处理附件
            attachments = []
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    
                    filename = part.get_filename()
                    if filename:
                        # 获取附件内容
                        attachment_data = part.get_payload(decode=True)
                        attachments.append({
                            'filename': filename,
                            'data': attachment_data,
                            'content_type': part.get_content_type()
                        })
                        logger.debug(f"获取附件: {filename}, 类型: {part.get_content_type()}, 大小: {len(attachment_data)} 字节")
            
            return attachments
        except Exception as e:
            logger.error(f"获取邮件附件失败: {e}")
            raise e
    
    def save_attachment(self, attachment, save_path=None):
        """保存附件到指定路径"""
        try:
            if save_path is None:
                import os
                # 默认保存到插件目录下的attachments文件夹
                save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachments")
                os.makedirs(save_path, exist_ok=True)
            
            file_path = os.path.join(save_path, attachment['filename'])
            with open(file_path, 'wb') as f:
                f.write(attachment['data'])
            
            logger.debug(f"附件已保存到: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"保存附件失败: {e}")
            raise e

    def query_mail(self, filter_keyword: str = None, filter_type: str = "UNSEEN"):
        """查询邮件"""
        try:
            import email
        except ImportError:
            raise Exception("email 模块未安装，请前往管理面板->控制台->安装pip库 安装 email 这个库")
        
        try:
            from email.header import decode_header
        except ImportError:
            raise Exception("email.header 模块未安装，请前往管理面板->控制台->安装pip库 安装 email 这个库")
        
        try:
            mail = self.login_mail()
                
            # 搜索未读邮件
            mail.select("INBOX")
            status, messages = mail.search(None, filter_type)
            mail_ids = messages[0].split()

            keywords = filter_keyword.split(",")
            mails = []
            for mail_id in mail_ids:
                _, msg_data = mail.fetch(mail_id, '(RFC822)')
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                if self.contains_keywords(msg, keywords):
                    # 处理附件
                    # attachments = self.get_mail_attachments(msg)

                    # 尝试解码邮件主题
                    subject = msg['Subject']
                    subject_str = " "
                    if subject:
                        try:
                            decoded_subject = decode_header(subject)
                            for text, charset in decoded_subject:
                                if isinstance(text, bytes):
                                    text = text.decode(charset or 'utf-8', errors='replace')
                                subject_str += text
                        except Exception as e:
                            pass
                    else:
                        pass

                    mails.append({
                        'subject': subject_str,
                        'from': msg['From'],
                        'to': msg['To'],
                        'date': msg['Date'],
                        'body': msg.get_payload(decode=True).decode(errors="ignore") if not msg.is_multipart() else None,
                        # 'attachments': attachments
                    })
            return mails
        except Exception as e:
            logger.debug(f"查询邮件失败: {e}")
            raise e

    @filter.command("mail_query")
    async def mail_query(self, event: AstrMessageEvent, filter_keyword: str = None, filter_type: str = "UNSEEN"):
        '''这是一个邮件查询指令'''
        # user_name = event.get_sender_name() # 用户名
        # message_str = event.message_str # 用户发的纯文本消息字符串
        # message_chain = event.get_messages() # 用户所发的消息的消息链 # from astrbot.api.message_components import *
        # logger.info(message_chain)
        # yield event.plain_result(f"Hello, {user_name}, 你发了 {message_str}!") # 发送一条纯文本消息
        mails = self.query_mail(filter_keyword, filter_type)
        yield event.plain_result(f"找到{len(mails)}封关键词中有{filter_keyword}的邮件")

    async def terminate(self):
        '''可选择实现 terminate 函数，当插件被卸载/停用时会调用。'''
        self.mail.close()
        self.mail.logout()
        self.mail = None

