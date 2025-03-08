import os
from astrbot.api.event import filter, AstrMessageEvent
from astrbot.api.star import Context, Star, register
from astrbot.api.message_components import Image, Plain
from astrbot.api import logger

@register("astrbot_plugin_mail", "mail", "一个邮件插件, 主要用于查询邮件", "1.1.7")
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
        if config["mail_address"] == "" or config["mail_password"] == "":
            raise Exception(
                "邮箱地址或密码未设置，请前往管理面板->插件管理->astrbot_plugin_mail 设置邮箱地址和密码"
            )
        else:
            self.mail = self.login_mail()
            self.test()

    def login_mail(self):
        """登录邮箱"""
        if self.mail is not None:
            return self.mail
        try:
            import imaplib
        except ImportError:
            raise Exception(
                "imaplib 模块未安装，请前往管理面板->控制台->安装pip库 安装 imaplib 这个库"
            )
        try:
            # 连接到 IMAP 服务器
            mail = imaplib.IMAP4_SSL(self.config["mail_host"], self.config["mail_port"])
            mail.login(
                self.config["mail_address"], self.config["mail_password"]
            )  # 使用您的 QQ 邮箱和授权码

            return mail
        except Exception as e:
            logger.error(f"登陆失败: {e}")
            raise e

    def get_mail_folders(self):
        # 获取所有文件夹
        status, folders = self.mail.list()
        print("可用文件夹：")
        for folder in folders:
            folder_name = folder.decode().split(' "/" ')[1].strip('"')
            print(folder_name)

        return folders

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

    def has_attachment(self, msg):
        """检查邮件是否包含附件"""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue
                filename = part.get_filename()
                if filename:
                    return True
        return False

    def get_mail_attachments(self, msg):
        """获取指定邮件的附件"""
        try:
            # 处理附件
            attachments = []
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue

                    filename = part.get_filename()
                    if filename:
                        # 解码文件名
                        if isinstance(filename, str):
                            # 如果已经是字符串，检查是否需要解码
                            try:
                                import email.header

                                decoded_header = email.header.decode_header(filename)
                                decoded_filename = ""
                                for text, charset in decoded_header:
                                    if isinstance(text, bytes):
                                        decoded_filename += text.decode(
                                            charset or "utf-8", errors="replace"
                                        )
                                    else:
                                        decoded_filename += text
                                filename = decoded_filename
                            except Exception as e:
                                logger.warning(f"解码文件名失败: {e}，使用原始文件名")

                        # 获取附件内容
                        attachment_data = part.get_payload(decode=True)
                        attachments.append(
                            {
                                "filename": filename,
                                "data": attachment_data,
                                "content_type": part.get_content_type(),
                            }
                        )

            return attachments
        except Exception as e:
            logger.error(f"获取邮件附件失败: {e}")
            raise e
        
    def get_attachment_path(self):
        """获取附件路径"""
        save_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "attachments"
        )
        if not os.path.exists(save_path):
            os.makedirs(save_path, exist_ok=True)
        return save_path
    
    def save_attachment(self, attachment):
        """保存附件到指定路径"""
        try:
            path = self.get_attachment_path()
            file_path = os.path.join(path, attachment["filename"])

            # 如果文件已存在，则直接输出
            if os.path.exists(file_path):
                return file_path

            with open(file_path, "wb") as f:
                f.write(attachment["data"])

            return file_path
        except Exception as e:
            logger.error(f"保存附件失败: {e}")
            raise e

    def parse_msg(self, msg, decode_header):
        # 尝试解码邮件主题
        subject = msg["Subject"]
        subject_str = " "
        if subject:
            try:
                decoded_subject = decode_header(subject)
                for text, charset in decoded_subject:
                    if isinstance(text, bytes):
                        text = text.decode(charset or "utf-8", errors="replace")
                    subject_str += text
            except Exception as e:
                pass
        else:
            pass

        # 获取邮件正文
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        # 解码收件人信息
        to_field = msg["To"]
        to_str = to_field
        if to_field:
            try:
                decoded_to = decode_header(to_field)
                to_str = ""
                for text, charset in decoded_to:
                    if isinstance(text, bytes):
                        text = text.decode(charset or "utf-8", errors="replace")
                    to_str += text
            except Exception as e:
                logger.error(f"解码收件人信息失败: {e}")

        # 解码发件人信息
        from_field = msg["From"]
        from_str = from_field
        if from_field:
            try:
                decoded_from = decode_header(from_field)
                from_str = ""
                for text, charset in decoded_from:
                    if isinstance(text, bytes):
                        text = text.decode(charset or "utf-8", errors="replace")
                    from_str += text
            except Exception as e:
                logger.error(f"解码发件人信息失败: {e}")

        return subject_str, from_str, to_str, body

    def query_mail(
        self,
        filter_keyword: str | None = None,
        filter_type: str = "UNSEEN",
        folder_name: str = "INBOX",
    ):
        """查询邮件"""
        try:
            import email
        except ImportError:
            raise Exception(
                "email 模块未安装，请前往管理面板->控制台->安装pip库 安装 email 这个库"
            )

        try:
            from email.header import decode_header
        except ImportError:
            raise Exception(
                "email.header 模块未安装，请前往管理面板->控制台->安装pip库 安装 email 这个库"
            )

        try:
            mail = self.login_mail()

            # 搜索未读邮件
            mail.select(folder_name)
            status, messages = mail.search(None, filter_type)
            mail_ids = messages[0].split()

            keywords = filter_keyword.split(",")
            mails = []
            for mail_id in mail_ids:
                _, msg_data = mail.fetch(mail_id, "(RFC822)")
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                if self.contains_keywords(msg, keywords):
                    subject_str, from_str, to_str, body = self.parse_msg(
                        msg, decode_header
                    )

                    mails.append(
                        {
                            "id": mail_id.decode('utf-8') if isinstance(mail_id, bytes) else str(mail_id),
                            "subject": subject_str,
                            "from": from_str,
                            "to": to_str,
                            "date": msg["Date"],
                            "body": body,
                            "has_attachment": self.has_attachment(msg),
                        }
                    )
            return mails
        except Exception as e:
            logger.debug(f"查询邮件失败: {e}")
            raise e

    def pdf_to_image(self, pdf_name: str, pdf_path: str):
        """将PDF转换为图片"""
        try:
            import fitz  # PyMuPDF
        except ImportError:
            raise Exception(
                "PyMuPDF 模块未安装，请前往管理面板->控制台->安装pip库 安装 PyMuPDF 这个库"
            )
        
        image_paths = []
        save_path = os.path.join(self.get_attachment_path())
        if not os.path.exists(save_path):
            os.makedirs(save_path, exist_ok=True)

        # 检查文件是否已存在，如果存在则直接返回
        if os.path.exists(save_path) and len(os.listdir(save_path)) > 0:
            # 获取已存在的图片路径
            for file_name in sorted(os.listdir(save_path)):
                if file_name.endswith('.png'):
                    image_paths.append(os.path.join(save_path, file_name))
            if image_paths:
                logger.debug(f"PDF已转换过图片，直接返回已有图片: {len(image_paths)}张")
                return image_paths
        
        pdf_document = fitz.open(pdf_path)
        for i, page in enumerate(pdf_document):
            pix = page.get_pixmap()
            image_path = os.path.join(save_path, f"{pdf_name}_{i}.png")
            pix.save(image_path)
            image_paths.append(image_path)
        pdf_document.close()

        return image_paths

    def get_attachment_file_by_id(
        self, mail_id: str | int, folder_name: str = "&UXZO1mWHTvZZOQ-/invoices"
    ):
        """根据邮件ID获取附件"""
        try:
            import email
        except ImportError:
            raise Exception(
                "email 模块未安装，请前往管理面板->控制台->安装pip库 安装 email 这个库"
            )

        try:
            mail = self.login_mail()
            mail.select(folder_name)

            if isinstance(mail_id, str):
                mail_id = mail_id.encode('utf-8')
            elif isinstance(mail_id, int):
                mail_id = str(mail_id).encode('utf-8')
            
            _, msg_data = mail.fetch(mail_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            files = []

            if self.has_attachment(msg):
                attachments = self.get_mail_attachments(msg)
                for attachment in attachments:
                    file_name = attachment["filename"]
                    file_path = self.save_attachment(attachment)
                    files.append({"file_name": file_name, "file_path": file_path})
                return files
            else:
                return files
        except Exception as e:
            logger.error(f"获取附件失败: {e}")
            raise e

    # 测试用
    def test(self):
        files = self.get_attachment_file_by_id('2', "&UXZO1mWHTvZZOQ-/invoices")
        imgs = []
        for file in files:
            imgs.extend(self.pdf_to_image(file["file_name"], file["file_path"]))
        logger.info(imgs)

        # logger.info(imgs)
        # self.get_mail_folders();
        # mails = self.query_mail("发票", "UNSEEN", "&UXZO1mWHTvZZOQ-/invoices")
        # reply_message = f"找到{len(mails)}封关键词中有发票的邮件\n"
        # for mail in mails:
        #     reply_message += f"----------------------------------\n"
        #     reply_message += f"ID: {mail['id']}\n"
        #     reply_message += f"主题: {mail['subject']}\n"
        #     reply_message += f"发件人: {mail['from']}\n"
        #     reply_message += f"日期: {mail['date']}\n"
        #     reply_message += f"正文: {mail['body']}\n"
        #     reply_message += (
        #         f"是否有附件: {'有' if mail['has_attachment'] else '没有'}\n"
        #     )
        #     reply_message += f"----------------------------------\n"
        # logger.info(reply_message)

    @filter.command("mail_query")
    async def mail_query(
        self,
        event: AstrMessageEvent,
        filter_keyword: str | None = None,
        filter_type: str = "UNSEEN",
        folder_name: str = "&UXZO1mWHTvZZOQ-/invoices",
    ):
        """这是一个邮件查询指令"""
        yield event.plain_result(f"查询关键词中有{filter_keyword}的邮件中，请稍后...")
        mails = self.query_mail(filter_keyword, filter_type, folder_name)
        reply_message = f"找到{len(mails)}封关键词中有{filter_keyword}的邮件\n"
        for mail in mails:
            reply_message += f"=========={mail['subject']}==========\n"
            reply_message += f"ID: {mail['id']}\n"
            reply_message += f"发件人: {mail['from']}\n"
            reply_message += f"日期: {mail['date']}\n"
            reply_message += f"正文: \n{mail['body']}\n"
            reply_message += (
                f"是否有附件: {'有' if mail['has_attachment'] else '没有'}\n"
            )
            reply_message += f"=========={mail['subject']}==========\n\n"
        yield event.plain_result(reply_message)

    @filter.command("mail_get_attachment")
    async def mail_get_attachment(
        self,
        event: AstrMessageEvent,
        mail_id: int | None = None,
        folder_name: str = "&UXZO1mWHTvZZOQ-/invoices",
    ):
        """获取邮件附件"""
        yield event.plain_result("正在获取附件，请稍后...")

        chain = []
        if mail_id is None:
            chain.append(Plain("请输入邮件ID"))
        else:
            files = self.get_attachment_file_by_id(mail_id, folder_name)
            if len(files) > 0:
                chain.append(Plain("附件如下："))
                for file in files:
                    chain.append(Plain(f"附件名称：{file['file_name']}，附件路径：{file['file_path']}"))
                    chain.append(Image.fromFileSystem(file["file_path"]))
            else:
                chain.append(Plain("没有找到附件"))

        yield event.chain_result(chain)
        
    async def terminate(self):
        """可选择实现 terminate 函数，当插件被卸载/停用时会调用。"""
        self.mail.close()
        self.mail.logout()
        self.mail = None
