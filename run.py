#!usr/bin/env python
# -*- coding:utf-8 _*-
"""
@author: caoping
@file:   run.py
@time:   2021/03/15
@desc:
"""
import os
import re
import json
import imaplib
import smtplib
import email
from datetime import datetime

# 根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open("_config.json") as f:
    config_str = f.read()

config = json.loads(config_str)


def temp_email_dir():
    """
    email 临时文件夹
    :return:
    """
    dir_name = datetime.strftime(datetime.today(), "%Y-%m-%d")
    temp_dir = os.path.join(BASE_DIR, dir_name)
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir)
    return temp_dir


def imap_server(config):
    # 接收邮件
    config_email = config['email']
    imapserver = imaplib.IMAP4_SSL(config_email['imapserver'], config_email['imapport'])
    imapserver.login(config_email['user'], config_email['password'])
    imapserver.select("INBOX")
    return imapserver


def parse_body(message, dir_path):
    for part in message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        # 附件
        filename = part.get_filename()
        if filename:
            # 附件内容
            file_data = part.get_payload(decode=True)
            # 邮件的路径
            file_path = os.path.join(dir_path, 'test.doc')
            # 二进制写入
            with open(file_path, 'wb') as f:
                f.write(file_data)


def get_email(imapserver):
    typ, data = imapserver.search(None, '(HEADER FROM "xxxx@sina.com")')
    msgList = data[0].split()
    latest = msgList[len(msgList) - 1]
    typ2, data2 = imapserver.fetch(latest, '(RFC822)')
    message = email.message_from_bytes(data2[0][1])
    return message


if __name__ == '__main__':
    imapserver = imap_server(config)
    message = get_email(imapserver)
    # 存放附件的路径
    temp_dir = temp_email_dir()
    parse_body(message, temp_dir)
