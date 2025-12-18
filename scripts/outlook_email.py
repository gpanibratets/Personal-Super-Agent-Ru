#!/usr/bin/env python3
"""
Скрипт для работы с Office 365 почтой через Microsoft Graph API.

Возможности:
- Чтение писем из почтового ящика
- Отправка писем
- Поиск писем
- Работа с вложениями

Требуется настройка:
1. Зарегистрировать приложение в Azure Portal
2. Получить Client ID и Client Secret
3. Настроить разрешения (Mail.Read, Mail.Send)
4. Сохранить данные в outlook_config.json
"""

import sys
import os
import json
from pathlib import Path
from datetime import datetime, timedelta

# Проверка наличия библиотеки O365
try:
    from O365 import Account
except ImportError:
    print("❌ Ошибка: Библиотека 'O365' не установлена.")
    print("   Установите её командой: pip3 install O365")
    sys.exit(1)

# Путь к конфигурационному файлу
CONFIG_FILE = Path(__file__).parent / "outlook_config.json"

def load_config():
    """Загружает конфигурацию из файла или создает шаблон."""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # Создаем шаблон конфигурации
        template = {
            "client_id": "YOUR_CLIENT_ID_HERE",
            "client_secret": "YOUR_CLIENT_SECRET_HERE",
            "tenant_id": "common",
            "scopes": ["basic", "message_all"],
            "email": "your_email@domain.com"
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(template, f, indent=2, ensure_ascii=False)
        print(f"⚠️  Создан файл конфигурации: {CONFIG_FILE}")
        print("📝 Пожалуйста, заполните client_id, client_secret и email в файле конфигурации.")
        print("\nИнструкция:")
        print("1. Зарегистрируйте приложение в Azure Portal (https://portal.azure.com)")
        print("2. Получите Client ID и Client Secret")
        print("3. Настройте разрешения: Mail.Read, Mail.Send")
        print("4. Заполните данные в outlook_config.json")
        print("\nПодробная инструкция: scripts/outlook_setup.md")
        return None

def get_account(config):
    """Создает и аутентифицирует аккаунт Office 365."""
    credentials = (config['client_id'], config['client_secret'])
    account = Account(credentials, tenant_id=config.get('tenant_id', 'common'))
    
    if account.authenticate(scopes=config.get('scopes', ['basic', 'message_all'])):
        return account
    else:
        print("❌ Ошибка аутентификации. Проверьте client_id и client_secret.")
        return None

def list_emails(account, limit=10, folder='inbox'):
    """Выводит список писем из почтового ящика."""
    mailbox = account.mailbox()
    
    if folder == 'inbox':
        inbox = mailbox.inbox_folder()
    elif folder == 'sent':
        inbox = mailbox.sent_folder()
    else:
        inbox = mailbox.inbox_folder()
    
    messages = inbox.get_messages(limit=limit, order_by='receivedDateTime desc')
    
    print(f"\n📧 Последние {limit} писем из папки '{folder}':\n")
    print(f"{'Дата':<20} {'От':<30} {'Тема':<50}")
    print("-" * 100)
    
    for message in messages:
        received = message.received.strftime('%Y-%m-%d %H:%M') if message.received else 'N/A'
        sender = message.sender.address if message.sender else 'N/A'
        subject = message.subject[:47] + '...' if len(message.subject) > 50 else message.subject
        
        print(f"{received:<20} {sender:<30} {subject:<50}")
    
    return messages

def read_email(account, message_id=None, index=0, folder='inbox'):
    """Читает конкретное письмо."""
    mailbox = account.mailbox()
    
    if folder == 'inbox':
        inbox = mailbox.inbox_folder()
    elif folder == 'sent':
        inbox = mailbox.sent_folder()
    else:
        inbox = mailbox.inbox_folder()
    
    if message_id:
        message = inbox.get_message(message_id)
    else:
        messages = inbox.get_messages(limit=index+1, order_by='receivedDateTime desc')
        if not messages or len(messages) <= index:
            print(f"❌ Письмо с индексом {index} не найдено.")
            return None
        message = messages[index]
    
    if not message:
        print("❌ Письмо не найдено.")
        return None
    
    print(f"\n📧 Письмо:")
    print(f"{'='*80}")
    print(f"От: {message.sender.address if message.sender else 'N/A'}")
    print(f"Кому: {', '.join([r.address for r in message.to]) if message.to else 'N/A'}")
    print(f"Тема: {message.subject}")
    print(f"Дата: {message.received.strftime('%Y-%m-%d %H:%M:%S') if message.received else 'N/A'}")
    print(f"{'='*80}")
    print(f"\n{message.body}")
    print(f"\n{'='*80}")
    
    # Вложения
    if message.attachments:
        print(f"\n📎 Вложения ({len(message.attachments)}):")
        for att in message.attachments:
            print(f"  - {att.name} ({att.size} bytes)")
    
    return message

def send_email(account, to, subject, body, attachments=None, cc=None, bcc=None):
    """Отправляет письмо."""
    mailbox = account.mailbox()
    message = mailbox.new_message()
    
    # Получатели
    if isinstance(to, str):
        message.to.add(to)
    else:
        for recipient in to:
            message.to.add(recipient)
    
    if cc:
        if isinstance(cc, str):
            message.cc.add(cc)
        else:
            for recipient in cc:
                message.cc.add(recipient)
    
    if bcc:
        if isinstance(bcc, str):
            message.bcc.add(bcc)
        else:
            for recipient in bcc:
                message.bcc.add(recipient)
    
    message.subject = subject
    message.body = body
    
    # Вложения
    if attachments:
        for att_path in attachments:
            if os.path.exists(att_path):
                message.attachments.add(att_path)
            else:
                print(f"⚠️  Вложение не найдено: {att_path}")
    
    if message.send():
        print(f"✅ Письмо успешно отправлено!")
        print(f"   Кому: {to}")
        print(f"   Тема: {subject}")
        return True
    else:
        print("❌ Ошибка при отправке письма.")
        return False

def search_emails(account, query, limit=10, folder='inbox'):
    """Ищет письма по запросу."""
    mailbox = account.mailbox()
    
    if folder == 'inbox':
        inbox = mailbox.inbox_folder()
    elif folder == 'sent':
        inbox = mailbox.sent_folder()
    else:
        inbox = mailbox.inbox_folder()
    
    # Поиск через фильтр
    messages = inbox.get_messages(limit=limit, query=query)
    
    print(f"\n🔍 Результаты поиска '{query}':\n")
    print(f"{'Дата':<20} {'От':<30} {'Тема':<50}")
    print("-" * 100)
    
    count = 0
    for message in messages:
        if query.lower() in message.subject.lower() or (message.body and query.lower() in message.body.lower()):
            received = message.received.strftime('%Y-%m-%d %H:%M') if message.received else 'N/A'
            sender = message.sender.address if message.sender else 'N/A'
            subject = message.subject[:47] + '...' if len(message.subject) > 50 else message.subject
            
            print(f"{received:<20} {sender:<30} {subject:<50}")
            count += 1
    
    if count == 0:
        print("Письма не найдены.")
    
    return messages

def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print(f"  {sys.argv[0]} list [--limit N] [--folder inbox|sent]")
        print(f"  {sys.argv[0]} read [--index N] [--id MESSAGE_ID] [--folder inbox|sent]")
        print(f"  {sys.argv[0]} send --to EMAIL --subject 'SUBJECT' --body 'BODY' [--attach FILE] [--cc EMAIL] [--bcc EMAIL]")
        print(f"  {sys.argv[0]} search --query 'QUERY' [--limit N] [--folder inbox|sent]")
        print("\nПримеры:")
        print(f"  {sys.argv[0]} list --limit 5")
        print(f"  {sys.argv[0]} read --index 0")
        print(f"  {sys.argv[0]} send --to 'user@example.com' --subject 'Test' --body 'Hello'")
        print(f"  {sys.argv[0]} search --query 'важно'")
        sys.exit(1)
    
    # Загружаем конфигурацию
    config = load_config()
    if not config:
        sys.exit(1)
    
    if config['client_id'] == "YOUR_CLIENT_ID_HERE" or config['client_secret'] == "YOUR_CLIENT_SECRET_HERE":
        print("❌ Ошибка: Необходимо настроить client_id и client_secret в файле конфигурации.")
        print(f"   Файл: {CONFIG_FILE}")
        sys.exit(1)
    
    # Аутентификация
    account = get_account(config)
    if not account:
        sys.exit(1)
    
    # Парсинг аргументов
    command = sys.argv[1]
    args = sys.argv[2:]
    
    # Обработка команд
    if command == 'list':
        limit = 10
        folder = 'inbox'
        
        if '--limit' in args:
            idx = args.index('--limit')
            if idx + 1 < len(args):
                limit = int(args[idx + 1])
        
        if '--folder' in args:
            idx = args.index('--folder')
            if idx + 1 < len(args):
                folder = args[idx + 1]
        
        list_emails(account, limit=limit, folder=folder)
    
    elif command == 'read':
        index = 0
        message_id = None
        folder = 'inbox'
        
        if '--index' in args:
            idx = args.index('--index')
            if idx + 1 < len(args):
                index = int(args[idx + 1])
        
        if '--id' in args:
            idx = args.index('--id')
            if idx + 1 < len(args):
                message_id = args[idx + 1]
        
        if '--folder' in args:
            idx = args.index('--folder')
            if idx + 1 < len(args):
                folder = args[idx + 1]
        
        read_email(account, message_id=message_id, index=index, folder=folder)
    
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
        
        send_email(account, to, subject, body, attachments=attachments, cc=cc, bcc=bcc)
    
    elif command == 'search':
        query = None
        limit = 10
        folder = 'inbox'
        
        if '--query' in args:
            idx = args.index('--query')
            if idx + 1 < len(args):
                query = args[idx + 1]
        
        if '--limit' in args:
            idx = args.index('--limit')
            if idx + 1 < len(args):
                limit = int(args[idx + 1])
        
        if '--folder' in args:
            idx = args.index('--folder')
            if idx + 1 < len(args):
                folder = args[idx + 1]
        
        if not query:
            print("❌ Ошибка: Укажите --query")
            sys.exit(1)
        
        search_emails(account, query, limit=limit, folder=folder)
    
    else:
        print(f"❌ Неизвестная команда: {command}")
        sys.exit(1)

if __name__ == "__main__":
    main()

