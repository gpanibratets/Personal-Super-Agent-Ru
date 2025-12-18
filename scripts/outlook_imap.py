#!/usr/bin/env python3
"""
Скрипт для работы с Office 365 почтой через IMAP/SMTP (без Azure).

Этот вариант не требует регистрации приложения в Azure и административных разрешений.
Использует стандартные протоколы IMAP для чтения и SMTP для отправки.

Требуется:
1. Включить IMAP в настройках Outlook
2. Создать пароль приложения (если включена двухфакторная аутентификация)
3. Настроить конфигурацию в outlook_imap_config.json
"""

import sys
import os
import json
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path
from datetime import datetime
import imaplib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Путь к конфигурационному файлу
CONFIG_FILE = Path(__file__).parent / "outlook_imap_config.json"

def load_config():
    """Загружает конфигурацию из файла или создает шаблон."""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # Создаем шаблон конфигурации
        template = {
            "email": "your_email@domain.com",
            "password": "YOUR_PASSWORD_OR_APP_PASSWORD_HERE",
            "imap_server": "outlook.office365.com",
            "imap_port": 993,
            "smtp_server": "smtp.office365.com",
            "smtp_port": 587
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(template, f, indent=2, ensure_ascii=False)
        print(f"⚠️  Создан файл конфигурации: {CONFIG_FILE}")
        print("📝 Пожалуйста, заполните email и password в файле конфигурации.")
        print("\nИнструкция:")
        print("1. Включите IMAP в настройках Outlook (если не включен)")
        print("2. Если включена двухфакторная аутентификация:")
        print("   - Создайте пароль приложения в настройках безопасности Microsoft")
        print("   - Используйте пароль приложения вместо обычного пароля")
        print("3. Заполните данные в outlook_imap_config.json")
        print("\nПодробная инструкция: scripts/outlook_imap_setup.md")
        return None

def decode_mime_words(s):
    """Декодирует MIME заголовки."""
    decoded_fragments = decode_header(s)
    decoded_str = ''
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            decoded_str += fragment.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_str += fragment
    return decoded_str

def list_emails(config, limit=10, folder='INBOX'):
    """Выводит список писем из почтового ящика."""
    try:
        # Подключение к IMAP серверу
        mail = imaplib.IMAP4_SSL(config['imap_server'], config['imap_port'])
        mail.login(config['email'], config['password'])
        mail.select(folder)
        
        # Поиск последних писем
        status, messages = mail.search(None, 'ALL')
        if status != 'OK':
            print("❌ Ошибка при поиске писем")
            return []
        
        email_ids = messages[0].split()
        email_ids = email_ids[-limit:] if len(email_ids) > limit else email_ids
        email_ids.reverse()  # Новые первыми
        
        print(f"\n📧 Последние {len(email_ids)} писем из папки '{folder}':\n")
        print(f"{'Дата':<20} {'От':<30} {'Тема':<50}")
        print("-" * 100)
        
        emails = []
        for email_id in email_ids:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            if status != 'OK':
                continue
            
            msg = email.message_from_bytes(msg_data[0][1])
            
            # Извлечение данных
            date_str = 'N/A'
            if msg['Date']:
                try:
                    date_obj = parsedate_to_datetime(msg['Date'])
                    date_str = date_obj.strftime('%Y-%m-%d %H:%M')
                except:
                    date_str = msg['Date'][:16]
            
            from_addr = decode_mime_words(msg['From'] or 'N/A')
            if len(from_addr) > 30:
                from_addr = from_addr[:27] + '...'
            
            subject = decode_mime_words(msg['Subject'] or '(без темы)')
            if len(subject) > 50:
                subject = subject[:47] + '...'
            
            print(f"{date_str:<20} {from_addr:<30} {subject:<50}")
            emails.append((email_id, msg))
        
        mail.close()
        mail.logout()
        return emails
        
    except imaplib.IMAP4.error as e:
        print(f"❌ Ошибка IMAP: {e}")
        print("💡 Проверьте:")
        print("   - Правильность email и пароля")
        print("   - Включен ли IMAP в настройках Outlook")
        print("   - Используете ли вы пароль приложения (если включена 2FA)")
        return []
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return []

def read_email(config, email_id=None, index=0, folder='INBOX'):
    """Читает конкретное письмо."""
    try:
        mail = imaplib.IMAP4_SSL(config['imap_server'], config['imap_port'])
        mail.login(config['email'], config['password'])
        mail.select(folder)
        
        if email_id:
            status, messages = mail.search(None, 'ALL')
            email_ids = messages[0].split()
            if email_id.encode() not in email_ids:
                print(f"❌ Письмо с ID {email_id} не найдено")
                mail.close()
                mail.logout()
                return None
            target_id = email_id.encode()
        else:
            status, messages = mail.search(None, 'ALL')
            email_ids = messages[0].split()
            email_ids.reverse()
            if len(email_ids) <= index:
                print(f"❌ Письмо с индексом {index} не найдено")
                mail.close()
                mail.logout()
                return None
            target_id = email_ids[index]
        
        status, msg_data = mail.fetch(target_id, '(RFC822)')
        if status != 'OK':
            print("❌ Ошибка при получении письма")
            mail.close()
            mail.logout()
            return None
        
        msg = email.message_from_bytes(msg_data[0][1])
        
        # Вывод информации о письме
        print(f"\n📧 Письмо:")
        print(f"{'='*80}")
        print(f"От: {decode_mime_words(msg['From'] or 'N/A')}")
        print(f"Кому: {decode_mime_words(msg['To'] or 'N/A')}")
        if msg['Cc']:
            print(f"Копия: {decode_mime_words(msg['Cc'])}")
        print(f"Тема: {decode_mime_words(msg['Subject'] or '(без темы)')}")
        print(f"Дата: {decode_mime_words(msg['Date'] or 'N/A')}")
        print(f"{'='*80}")
        
        # Извлечение тела письма
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                        break
                    except:
                        pass
                elif content_type == "text/html" and not body:
                    try:
                        body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        pass
        else:
            try:
                body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
            except:
                body = str(msg.get_payload())
        
        print(f"\n{body}")
        print(f"\n{'='*80}")
        
        # Вложения
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        filename = decode_mime_words(filename)
                        attachments.append(filename)
        
        if attachments:
            print(f"\n📎 Вложения ({len(attachments)}):")
            for att in attachments:
                print(f"  - {att}")
        
        mail.close()
        mail.logout()
        return msg
        
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return None

def send_email(config, to, subject, body, attachments=None, cc=None, bcc=None):
    """Отправляет письмо через SMTP."""
    try:
        # Создание сообщения
        msg = MIMEMultipart()
        msg['From'] = config['email']
        msg['To'] = to if isinstance(to, str) else ', '.join(to)
        if cc:
            msg['Cc'] = cc if isinstance(cc, str) else ', '.join(cc)
        msg['Subject'] = subject
        
        # Тело письма
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Вложения
        if attachments:
            for file_path in attachments:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {os.path.basename(file_path)}'
                    )
                    msg.attach(part)
                else:
                    print(f"⚠️  Вложение не найдено: {file_path}")
        
        # Отправка
        server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
        server.starttls()
        server.login(config['email'], config['password'])
        
        recipients = [to] if isinstance(to, str) else to
        if cc:
            recipients.extend([cc] if isinstance(cc, str) else cc)
        if bcc:
            recipients.extend([bcc] if isinstance(bcc, str) else bcc)
        
        text = msg.as_string()
        server.sendmail(config['email'], recipients, text)
        server.quit()
        
        print(f"✅ Письмо успешно отправлено!")
        print(f"   Кому: {to}")
        print(f"   Тема: {subject}")
        return True
        
    except smtplib.SMTPAuthenticationError:
        print("❌ Ошибка аутентификации SMTP")
        print("💡 Проверьте:")
        print("   - Правильность email и пароля")
        print("   - Используете ли вы пароль приложения (если включена 2FA)")
        return False
    except Exception as e:
        print(f"❌ Ошибка при отправке: {e}")
        return False

def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print(f"  {sys.argv[0]} list [--limit N] [--folder INBOX|Sent]")
        print(f"  {sys.argv[0]} read [--index N] [--folder INBOX|Sent]")
        print(f"  {sys.argv[0]} send --to EMAIL --subject 'SUBJECT' --body 'BODY' [--attach FILE] [--cc EMAIL] [--bcc EMAIL]")
        print("\nПримеры:")
        print(f"  {sys.argv[0]} list --limit 5")
        print(f"  {sys.argv[0]} read --index 0")
        print(f"  {sys.argv[0]} send --to 'user@example.com' --subject 'Test' --body 'Hello'")
        sys.exit(1)
    
    # Загружаем конфигурацию
    config = load_config()
    if not config:
        sys.exit(1)
    
    if config['email'] == "your_email@domain.com" or config['password'] == "YOUR_PASSWORD_OR_APP_PASSWORD_HERE":
        print("❌ Ошибка: Необходимо настроить email и password в файле конфигурации.")
        print(f"   Файл: {CONFIG_FILE}")
        sys.exit(1)
    
    # Парсинг аргументов
    command = sys.argv[1]
    args = sys.argv[2:]
    
    # Обработка команд
    if command == 'list':
        limit = 10
        folder = 'INBOX'
        
        if '--limit' in args:
            idx = args.index('--limit')
            if idx + 1 < len(args):
                limit = int(args[idx + 1])
        
        if '--folder' in args:
            idx = args.index('--folder')
            if idx + 1 < len(args):
                folder = args[idx + 1]
        
        list_emails(config, limit=limit, folder=folder)
    
    elif command == 'read':
        index = 0
        folder = 'INBOX'
        
        if '--index' in args:
            idx = args.index('--index')
            if idx + 1 < len(args):
                index = int(args[idx + 1])
        
        if '--folder' in args:
            idx = args.index('--folder')
            if idx + 1 < len(args):
                folder = args[idx + 1]
        
        read_email(config, index=index, folder=folder)
    
    elif command == 'send':
        to = None
        subject = None
        body = None
        attachments = []
        cc = None
        bcc = None
        
        if '--to' in args:
            idx = args.index('--to')
            if idx + 1 < len(args):
                to = args[idx + 1]
        
        if '--subject' in args:
            idx = args.index('--subject')
            if idx + 1 < len(args):
                subject = args[idx + 1]
        
        if '--body' in args:
            idx = args.index('--body')
            if idx + 1 < len(args):
                body = args[idx + 1]
        
        if '--attach' in args:
            idx = args.index('--attach')
            while idx + 1 < len(args) and not args[idx + 1].startswith('--'):
                attachments.append(args[idx + 1])
                idx += 1
        
        if '--cc' in args:
            idx = args.index('--cc')
            if idx + 1 < len(args):
                cc = args[idx + 1]
        
        if '--bcc' in args:
            idx = args.index('--bcc')
            if idx + 1 < len(args):
                bcc = args[idx + 1]
        
        if not to or not subject or not body:
            print("❌ Ошибка: Укажите --to, --subject и --body")
            sys.exit(1)
        
        send_email(config, to, subject, body, attachments=attachments, cc=cc, bcc=bcc)
    
    else:
        print(f"❌ Неизвестная команда: {command}")
        sys.exit(1)

if __name__ == "__main__":
    main()

