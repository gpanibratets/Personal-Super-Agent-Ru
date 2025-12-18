#!/usr/bin/env python3
"""
Скрипт для работы с Office 365 почтой через Exchange Web Services (EWS).

Этот вариант использует протокол Exchange и не требует Azure регистрации.
Работает напрямую с Exchange сервером Office 365.

Требуется:
1. Email адрес Office 365
2. Пароль или пароль приложения (если включена 2FA)
3. URL Exchange сервера (обычно определяется автоматически)
"""

import sys
import os
import json
from pathlib import Path
from datetime import datetime, timedelta

# Проверка наличия библиотеки exchangelib
try:
    from exchangelib import Credentials, Account, Message, Mailbox, FileAttachment
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    import requests
    from requests.adapters import HTTPAdapter
except ImportError:
    print("❌ Ошибка: Библиотека 'exchangelib' не установлена.")
    print("   Установите её командой: pip3 install exchangelib")
    sys.exit(1)

# Отключение проверки SSL (может потребоваться для корпоративных серверов)
# Будет включено через конфигурацию если нужно

# Путь к конфигурационному файлу
CONFIG_FILE = Path(__file__).parent / "outlook_exchange_config.json"

def load_config():
    """Загружает конфигурацию из файла или создает шаблон."""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # Создаем шаблон конфигурации
        template = {
            "email": "your_email@domain.com",
            "username": null,
            "password": "YOUR_PASSWORD_OR_APP_PASSWORD_HERE",
            "server": "outlook.office365.com",
            "autodiscover": True,
            "verify_ssl": True,
            "verify_ssl": True
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(template, f, indent=2, ensure_ascii=False)
        print(f"⚠️  Создан файл конфигурации: {CONFIG_FILE}")
        print("📝 Пожалуйста, заполните email и password в файле конфигурации.")
        print("\nИнструкция:")
        print("1. Если включена двухфакторная аутентификация:")
        print("   - Создайте пароль приложения в настройках безопасности Microsoft")
        print("   - Используйте пароль приложения вместо обычного пароля")
        print("2. Заполните данные в outlook_exchange_config.json")
        print("\nПодробная инструкция: scripts/outlook_exchange_setup.md")
        return None

def get_account(config):
    """Создает и подключается к Exchange аккаунту."""
    try:
        # Настройка SSL проверки
        verify_ssl = config.get('verify_ssl', True)
        if not verify_ssl:
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            print("⚠️  Проверка SSL сертификата отключена")
        
        # Подготовка учетных данных
        email = config['email']
        password = config['password']
        username = config.get('username')  # Формат DOMAIN\username или просто username
        
        # Если указан username (в формате DOMAIN\username), используем его
        if username:
            # Заменяем двойной обратный слэш на одинарный (если был экранирован в JSON)
            username = username.replace('\\\\', '\\')
            print(f"🔐 Использование логина: {username}")
            
            # Для корпоративных серверов используем базовую аутентификацию с доменным логином
            # exchangelib автоматически определит нужный метод аутентификации
            try:
                credentials = Credentials(username, password)
            except Exception as e:
                print(f"⚠️  Ошибка создания credentials: {e}")
                # Пробуем использовать email как fallback
                print(f"   Пробую использовать email: {email}")
                credentials = Credentials(email, password)
        else:
            # Иначе используем email
            print(f"🔐 Использование email: {email}")
            credentials = Credentials(email, password)
        
        # Если указан сервер и autodiscover отключен, используем указанный сервер
        server = config.get('server', '').strip()
        use_autodiscover = config.get('autodiscover', True)
        
        if server and server != 'outlook.office365.com' and not use_autodiscover:
            # Использование указанного сервера
            from exchangelib import Configuration
            # Если указан полный URL, извлекаем только имя сервера
            if server.startswith('http'):
                from urllib.parse import urlparse
                parsed = urlparse(server)
                server = parsed.hostname or server.split('://')[1].split('/')[0]
            # Убираем порт если есть
            if ':' in server:
                server = server.split(':')[0]
            print(f"🔗 Подключение к Exchange серверу: {server}")
            
            config_exchange = Configuration(server=server, credentials=credentials)
            # Для корпоративных серверов используем access_type='delegate'
            # Account автоматически определит правильный email из credentials
            account = Account(email, config=config_exchange, access_type='delegate')
        else:
            # Автоматическое определение сервера
            print("🔍 Автоматическое определение Exchange сервера...")
            account = Account(config['email'], credentials=credentials, autodiscover=True)
        
        return account
    except Exception as e:
        error_msg = str(e)
        print(f"❌ Ошибка подключения к Exchange: {error_msg}")
        print("\n💡 Возможные решения:")
        print("   1. Проверьте правильность email и пароля")
        print("   2. Если включена 2FA, используйте пароль приложения")
        print("   3. Для корпоративных серверов может потребоваться:")
        print("      - Доменное имя в формате DOMAIN\\username")
        print("      - Полный email адрес как username")
        print("      - NTLM аутентификация")
        print("   4. Проверьте доступность Exchange протокола для вашего аккаунта")
        print("\n📝 Попробуйте добавить в конфигурацию:")
        print('   "domain": "ваш_домен"  (если требуется доменная аутентификация)')
        return None

def list_emails(account, limit=10, folder='inbox'):
    """Выводит список писем из почтового ящика."""
    try:
        if folder == 'inbox':
            mailbox = account.inbox
        elif folder == 'sent':
            mailbox = account.sent
        elif folder == 'drafts':
            mailbox = account.drafts
        else:
            mailbox = account.inbox
        
        # Получение писем, отсортированных по дате получения (новые первыми)
        items = mailbox.all().order_by('-datetime_received')[:limit]
        
        print(f"\n📧 Последние {limit} писем из папки '{folder}':\n")
        print(f"{'Дата':<20} {'От':<30} {'Тема':<50}")
        print("-" * 100)
        
        emails = []
        for item in items:
            date_str = item.datetime_received.strftime('%Y-%m-%d %H:%M') if item.datetime_received else 'N/A'
            sender = item.sender.email_address if item.sender else 'N/A'
            if len(sender) > 30:
                sender = sender[:27] + '...'
            
            subject = item.subject[:47] + '...' if len(item.subject) > 50 else (item.subject or '(без темы)')
            
            print(f"{date_str:<20} {sender:<30} {subject:<50}")
            emails.append(item)
        
        return emails
        
    except Exception as e:
        print(f"❌ Ошибка при получении писем: {e}")
        return []

def read_email(account, email_id=None, index=0, folder='inbox'):
    """Читает конкретное письмо."""
    try:
        if folder == 'inbox':
            mailbox = account.inbox
        elif folder == 'sent':
            mailbox = account.sent
        elif folder == 'drafts':
            mailbox = account.drafts
        else:
            mailbox = account.inbox
        
        if email_id:
            # Поиск по ID
            item = mailbox.get(id=email_id)
        else:
            # Получение по индексу
            items = list(mailbox.all().order_by('-datetime_received')[:index+1])
            if not items or len(items) <= index:
                print(f"❌ Письмо с индексом {index} не найдено.")
                return None
            item = items[index]
        
        # Вывод информации о письме
        print(f"\n📧 Письмо:")
        print(f"{'='*80}")
        print(f"От: {item.sender.email_address if item.sender else 'N/A'}")
        print(f"Кому: {', '.join([r.email_address for r in item.to_recipients]) if item.to_recipients else 'N/A'}")
        if item.cc_recipients:
            print(f"Копия: {', '.join([r.email_address for r in item.cc_recipients])}")
        print(f"Тема: {item.subject or '(без темы)'}")
        print(f"Дата: {item.datetime_received.strftime('%Y-%m-%d %H:%M:%S') if item.datetime_received else 'N/A'}")
        print(f"{'='*80}")
        
        # Тело письма
        body = item.body or item.text_body or ''
        if hasattr(body, 'strip'):
            print(f"\n{body.strip()}")
        else:
            print(f"\n{str(body)}")
        print(f"\n{'='*80}")
        
        # Вложения
        if item.attachments:
            print(f"\n📎 Вложения ({len(item.attachments)}):")
            for att in item.attachments:
                if isinstance(att, FileAttachment):
                    print(f"  - {att.name} ({att.size} bytes)")
        
        return item
        
    except Exception as e:
        print(f"❌ Ошибка при чтении письма: {e}")
        return None

def send_email(account, to, subject, body, attachments=None, cc=None, bcc=None):
    """Отправляет письмо через Exchange."""
    try:
        # Создание сообщения
        m = Message(
            account=account,
            subject=subject,
            body=body,
            to_recipients=[Mailbox(email_address=to)] if isinstance(to, str) else [Mailbox(email_address=addr) for addr in to]
        )
        
        if cc:
            m.cc_recipients = [Mailbox(email_address=cc)] if isinstance(cc, str) else [Mailbox(email_address=addr) for addr in cc]
        
        if bcc:
            m.bcc_recipients = [Mailbox(email_address=bcc)] if isinstance(bcc, str) else [Mailbox(email_address=addr) for addr in bcc]
        
        # Вложения
        if attachments:
            for file_path in attachments:
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        file_content = f.read()
                    att = FileAttachment(name=os.path.basename(file_path), content=file_content)
                    m.attachments.append(att)
                else:
                    print(f"⚠️  Вложение не найдено: {file_path}")
        
        # Отправка
        m.send()
        
        print(f"✅ Письмо успешно отправлено!")
        print(f"   Кому: {to}")
        print(f"   Тема: {subject}")
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при отправке письма: {e}")
        return False

def search_emails(account, query, limit=10, folder='inbox'):
    """Ищет письма по запросу."""
    try:
        if folder == 'inbox':
            mailbox = account.inbox
        elif folder == 'sent':
            mailbox = account.sent
        else:
            mailbox = account.inbox
        
        # Поиск по теме и телу письма
        items = mailbox.filter(
            subject__contains=query
        ) | mailbox.filter(
            body__contains=query
        )
        
        items = items.order_by('-datetime_received')[:limit]
        
        print(f"\n🔍 Результаты поиска '{query}':\n")
        print(f"{'Дата':<20} {'От':<30} {'Тема':<50}")
        print("-" * 100)
        
        count = 0
        for item in items:
            date_str = item.datetime_received.strftime('%Y-%m-%d %H:%M') if item.datetime_received else 'N/A'
            sender = item.sender.email_address if item.sender else 'N/A'
            if len(sender) > 30:
                sender = sender[:27] + '...'
            
            subject = item.subject[:47] + '...' if len(item.subject) > 50 else (item.subject or '(без темы)')
            
            print(f"{date_str:<20} {sender:<30} {subject:<50}")
            count += 1
        
        if count == 0:
            print("Письма не найдены.")
        
        return list(items)
        
    except Exception as e:
        print(f"❌ Ошибка при поиске: {e}")
        return []

def test_connection(config):
    """Тестирует подключение с разными вариантами аутентификации."""
    print("\n🔍 Тестирование подключения...\n")
    
    variants = []
    if config.get('username'):
        variants.append(('username', config['username']))
        # Пробуем только число без домена
        if '\\' in config['username']:
            variants.append(('username (только число)', config['username'].split('\\')[-1]))
    variants.append(('email', config['email']))
    
    for name, login in variants:
        print(f"Пробую: {name} = {login}")
        try:
            credentials = Credentials(login, config['password'])
            from exchangelib import Configuration
            config_exchange = Configuration(server=config['server'], credentials=credentials)
            # Пробуем разные варианты email для Account
            account_emails = [config['email']]
            if config.get('username'):
                account_emails.append(config['username'].split('\\')[-1] if '\\' in config['username'] else config['username'])
            
            for acc_email in account_emails:
                try:
                    account = Account(acc_email, config=config_exchange)
                    # Пробуем получить inbox
                    inbox = account.inbox
                    items = list(inbox.all().order_by('-datetime_received')[:1])
                    print(f"✅ Успешно! Использован email: {acc_email}, найдено писем: {len(items)}")
                    return account
                except Exception as e2:
                    if acc_email != account_emails[-1]:
                        continue
                    raise e2
        except Exception as e:
            error_msg = str(e)
            if 'credentials' in error_msg.lower() or 'authentication' in error_msg.lower():
                print(f"   ❌ Ошибка аутентификации")
            else:
                print(f"   ❌ Ошибка: {error_msg[:80]}")
    
    return None

def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print(f"  {sys.argv[0]} test  # тестирование подключения с разными вариантами")
        print(f"  {sys.argv[0]} list [--limit N] [--folder inbox|sent|drafts]")
        print(f"  {sys.argv[0]} read [--index N] [--id EMAIL_ID] [--folder inbox|sent|drafts]")
        print(f"  {sys.argv[0]} send --to EMAIL --subject 'SUBJECT' --body 'BODY' [--attach FILE] [--cc EMAIL] [--bcc EMAIL]")
        print(f"  {sys.argv[0]} search --query 'QUERY' [--limit N] [--folder inbox|sent]")
        print("\nПримеры:")
        print(f"  {sys.argv[0]} test  # протестировать подключение")
        print(f"  {sys.argv[0]} list --limit 5")
        print(f"  {sys.argv[0]} read --index 0")
        print(f"  {sys.argv[0]} send --to 'user@example.com' --subject 'Test' --body 'Hello'")
        print(f"  {sys.argv[0]} search --query 'важно'")
        sys.exit(1)
    
    # Загружаем конфигурацию
    config = load_config()
    if not config:
        sys.exit(1)
    
    if config['email'] == "your_email@domain.com" or config['password'] == "YOUR_PASSWORD_OR_APP_PASSWORD_HERE":
        print("❌ Ошибка: Необходимо настроить email и password в файле конфигурации.")
        print(f"   Файл: {CONFIG_FILE}")
        sys.exit(1)
    
    # Подключение к Exchange
    account = get_account(config)
    if not account:
        sys.exit(1)
    
    # Парсинг аргументов
    command = sys.argv[1]
    args = sys.argv[2:]
    
    # Специальная команда для тестирования
    if command == 'test':
        account = test_connection(config)
        if account:
            print("\n✅ Подключение работает! Можно использовать команды list, read, send")
        sys.exit(0)
    
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
        email_id = None
        folder = 'inbox'
        
        if '--index' in args:
            idx = args.index('--index')
            if idx + 1 < len(args):
                index = int(args[idx + 1])
        
        if '--id' in args:
            idx = args.index('--id')
            if idx + 1 < len(args):
                email_id = args[idx + 1]
        
        if '--folder' in args:
            idx = args.index('--folder')
            if idx + 1 < len(args):
                folder = args[idx + 1]
        
        read_email(account, email_id=email_id, index=index, folder=folder)
    
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

