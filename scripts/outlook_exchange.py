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

# Для работы с часовыми поясами
try:
    from zoneinfo import ZoneInfo
except ImportError:
    # Для Python < 3.9 используем pytz
    try:
        import pytz
        ZoneInfo = pytz.timezone
    except ImportError:
        print("⚠️  Для работы с часовыми поясами установите pytz: pip3 install pytz")
        ZoneInfo = None

# Проверка наличия библиотеки exchangelib
try:
    from exchangelib import Credentials, Account, Message, Mailbox, FileAttachment
    from exchangelib import CalendarItem, EWSDateTime, EWSTimeZone
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    from exchangelib.folders import Calendar
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

def convert_to_almaty_time(ews_datetime):
    """Конвертирует EWSDateTime в часовой пояс Алматы (UTC+6)."""
    if not ews_datetime:
        return None
    
    try:
        # Получаем часовой пояс Алматы
        almaty_tz = ZoneInfo('Asia/Almaty')
        
        # EWSDateTime наследуется от datetime и имеет метод astimezone
        # Просто используем его напрямую - это самый надежный способ
        if hasattr(ews_datetime, 'astimezone'):
            try:
                dt = ews_datetime.astimezone(almaty_tz)
                return dt
            except Exception as e:
                # Если astimezone не работает, возвращаем исходное время
                return ews_datetime
        
        return ews_datetime
    except Exception as e:
        # Если конвертация не удалась, возвращаем исходное время
        return ews_datetime

def parse_datetime(date_str, default_timezone=None):
    """Парсит строку даты в EWSDateTime."""
    try:
        # Пробуем разные форматы
        formats = [
            '%Y-%m-%d %H:%M',
            '%Y-%m-%dT%H:%M:%S',
            '%Y-%m-%dT%H:%M:%SZ',
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d'
        ]
        
        for fmt in formats:
            try:
                dt = datetime.strptime(date_str, fmt)
                # Если нет часового пояса, используем переданный или UTC
                if default_timezone:
                    # Создаем timezone-aware datetime
                    dt_aware = dt.replace(tzinfo=default_timezone)
                    return EWSDateTime.from_datetime(dt_aware)
                else:
                    # Используем UTC по умолчанию
                    from exchangelib import UTC
                    dt_aware = dt.replace(tzinfo=UTC)
                    return EWSDateTime.from_datetime(dt_aware)
            except ValueError:
                continue
        
        # Если не получилось, пробуем from_string (для ISO форматов)
        try:
            return EWSDateTime.from_string(date_str)
        except:
            pass
        
        # Если все не удалось, возвращаем текущее время
        if default_timezone:
            return EWSDateTime.now(tz=default_timezone)
        else:
            return EWSDateTime.now()
    except Exception as e:
        print(f"⚠️  Ошибка парсинга даты '{date_str}': {e}")
        if default_timezone:
            return EWSDateTime.now(tz=default_timezone)
        else:
            return EWSDateTime.now()

def list_calendar(account, limit=10, start_date=None, end_date=None):
    """Получает список событий календаря."""
    try:
        calendar = account.calendar
        tz = account.default_timezone
        
        # Определение диапазона дат
        if start_date:
            if isinstance(start_date, str):
                start = parse_datetime(start_date, tz)
            else:
                start = start_date
        else:
            # По умолчанию - сегодня
            start = EWSDateTime.now(tz=tz)
        
        if end_date:
            if isinstance(end_date, str):
                end = parse_datetime(end_date, tz)
            else:
                end = end_date
        else:
            # По умолчанию - через 30 дней
            end = start + timedelta(days=30)
        
        # Получение событий
        items = calendar.view(
            start=start,
            end=end
        ).order_by('start')[:limit]
        
        print(f"\n📅 События календаря ({start.date()} - {end.date()}):\n")
        print(f"{'Дата/Время':<25} {'Тема':<50} {'Участники':<30}")
        print("-" * 105)
        
        count = 0
        for item in items:
            # Конвертируем время в часовой пояс Алматы
            if item.start:
                try:
                    # Получаем часовой пояс Алматы
                    almaty_tz = ZoneInfo('Asia/Almaty')
                    # EWSDateTime всегда timezone-aware, конвертируем напрямую
                    almaty_time = item.start.astimezone(almaty_tz)
                    start_str = almaty_time.strftime('%Y-%m-%d %H:%M')
                except Exception as e:
                    # Если конвертация вызвала ошибку, используем исходное время
                    start_str = item.start.strftime('%Y-%m-%d %H:%M') if hasattr(item.start, 'strftime') else str(item.start)
            else:
                start_str = 'N/A'
            
            subject = (item.subject[:47] + '...') if item.subject and len(item.subject) > 50 else (item.subject or '(без темы)')
            
            # Участники
            attendees = []
            if hasattr(item, 'required_attendees') and item.required_attendees:
                attendees.extend([a.mailbox.email_address for a in item.required_attendees if a.mailbox])
            if hasattr(item, 'optional_attendees') and item.optional_attendees:
                attendees.extend([a.mailbox.email_address for a in item.optional_attendees if a.mailbox])
            attendees_str = ', '.join(attendees[:2]) if attendees else 'Нет участников'
            if len(attendees) > 2:
                attendees_str += f' (+{len(attendees)-2})'
            if len(attendees_str) > 30:
                attendees_str = attendees_str[:27] + '...'
            
            print(f"{start_str:<25} {subject:<50} {attendees_str:<30}")
            count += 1
        
        if count == 0:
            print("События не найдены.")
        
        return list(items)
        
    except Exception as e:
        print(f"❌ Ошибка при получении календаря: {e}")
        return []

def create_meeting(account, subject, start_time, end_time, attendees=None, body=None, location=None):
    """Создает встречу в календаре."""
    try:
        tz = account.default_timezone
        
        # Преобразование времени
        if isinstance(start_time, str):
            start = parse_datetime(start_time, tz)
        else:
            start = start_time
        
        if isinstance(end_time, str):
            end = parse_datetime(end_time, tz)
        else:
            end = end_time
        
        # Создание встречи
        meeting = CalendarItem(
            account=account,
            folder=account.calendar,
            subject=subject,
            start=start,
            end=end,
            body=body or '',
            location=location or '',
            required_attendees=[Mailbox(email_address=email) for email in attendees] if attendees else []
        )
        
        # Сохранение и отправка приглашений
        meeting.save(send_meeting_invitations='SendToAllAndSaveCopy')
        
        print(f"✅ Встреча успешно создана!")
        print(f"   Тема: {subject}")
        print(f"   Время: {start.strftime('%Y-%m-%d %H:%M')} - {end.strftime('%Y-%m-%d %H:%M')}")
        if attendees:
            print(f"   Участники: {', '.join(attendees)}")
        if location:
            print(f"   Место: {location}")
        
        return meeting
        
    except Exception as e:
        print(f"❌ Ошибка при создании встречи: {e}")
        return None

def search_calendar(account, query, limit=10, start_date=None, end_date=None):
    """Ищет события в календаре по запросу."""
    try:
        calendar = account.calendar
        tz = account.default_timezone
        
        # Определение диапазона дат
        if start_date:
            if isinstance(start_date, str):
                start = parse_datetime(start_date, tz)
            else:
                start = start_date
        else:
            start = EWSDateTime.now(tz=tz)
        
        if end_date:
            if isinstance(end_date, str):
                end = parse_datetime(end_date, tz)
            else:
                end = end_date
        else:
            end = start + timedelta(days=365)  # Год вперед
        
        # Поиск по теме (основной поиск)
        items = calendar.filter(
            start__gte=start,
            start__lte=end,
            subject__contains=query
        ).order_by('start')[:limit]
        
        print(f"\n🔍 Результаты поиска '{query}':\n")
        print(f"{'Дата/Время':<25} {'Тема':<50}")
        print("-" * 75)
        
        count = 0
        for item in items:
            # Конвертируем время в часовой пояс Алматы
            if item.start:
                try:
                    # Получаем часовой пояс Алматы
                    almaty_tz = ZoneInfo('Asia/Almaty')
                    # EWSDateTime всегда timezone-aware, конвертируем напрямую
                    almaty_time = item.start.astimezone(almaty_tz)
                    start_str = almaty_time.strftime('%Y-%m-%d %H:%M')
                except Exception as e:
                    # Если конвертация вызвала ошибку, используем исходное время
                    start_str = item.start.strftime('%Y-%m-%d %H:%M') if hasattr(item.start, 'strftime') else str(item.start)
            else:
                start_str = 'N/A'
            
            subject = (item.subject[:47] + '...') if item.subject and len(item.subject) > 50 else (item.subject or '(без темы)')
            print(f"{start_str:<25} {subject:<50}")
            count += 1
        
        if count == 0:
            print("События не найдены.")
        
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
        print(f"\n📅 Календарь:")
        print(f"  {sys.argv[0]} calendar [--limit N] [--start DATE] [--end DATE]")
        print(f"  {sys.argv[0]} calendar-create --subject 'SUBJECT' --start 'YYYY-MM-DD HH:MM' --end 'YYYY-MM-DD HH:MM' [--attendees EMAIL1,EMAIL2] [--body 'BODY'] [--location 'LOCATION']")
        print(f"  {sys.argv[0]} calendar-search --query 'QUERY' [--limit N] [--start DATE] [--end DATE]")
        print("\nПримеры:")
        print(f"  {sys.argv[0]} test  # протестировать подключение")
        print(f"  {sys.argv[0]} list --limit 5")
        print(f"  {sys.argv[0]} read --index 0")
        print(f"  {sys.argv[0]} send --to 'user@example.com' --subject 'Test' --body 'Hello'")
        print(f"  {sys.argv[0]} search --query 'важно'")
        print(f"  {sys.argv[0]} calendar --limit 10")
        print(f"  {sys.argv[0]} calendar-create --subject 'Встреча' --start '2025-12-24 09:00' --end '2025-12-24 10:00' --attendees 'user@example.com'")
        print(f"  {sys.argv[0]} calendar-search --query 'Profitbase'")
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
    
    elif command == 'calendar':
        limit = 10
        start_date = None
        end_date = None
        
        if '--limit' in args:
            idx = args.index('--limit')
            if idx + 1 < len(args):
                limit = int(args[idx + 1])
        
        if '--start' in args:
            idx = args.index('--start')
            if idx + 1 < len(args):
                start_date = args[idx + 1]
        
        if '--end' in args:
            idx = args.index('--end')
            if idx + 1 < len(args):
                end_date = args[idx + 1]
        
        list_calendar(account, limit=limit, start_date=start_date, end_date=end_date)
    
    elif command == 'calendar-create':
        subject = None
        start_time = None
        end_time = None
        attendees = None
        body = None
        location = None
        
        if '--subject' in args:
            idx = args.index('--subject')
            if idx + 1 < len(args):
                subject = args[idx + 1]
        
        if '--start' in args:
            idx = args.index('--start')
            if idx + 1 < len(args):
                start_time = args[idx + 1]
        
        if '--end' in args:
            idx = args.index('--end')
            if idx + 1 < len(args):
                end_time = args[idx + 1]
        
        if '--attendees' in args:
            idx = args.index('--attendees')
            if idx + 1 < len(args):
                attendees = [email.strip() for email in args[idx + 1].split(',')]
        
        if '--body' in args:
            idx = args.index('--body')
            if idx + 1 < len(args):
                body = args[idx + 1]
        
        if '--location' in args:
            idx = args.index('--location')
            if idx + 1 < len(args):
                location = args[idx + 1]
        
        if not subject or not start_time or not end_time:
            print("❌ Ошибка: Укажите --subject, --start и --end")
            sys.exit(1)
        
        create_meeting(account, subject, start_time, end_time, attendees=attendees, body=body, location=location)
    
    elif command == 'calendar-search':
        query = None
        limit = 10
        start_date = None
        end_date = None
        
        if '--query' in args:
            idx = args.index('--query')
            if idx + 1 < len(args):
                query = args[idx + 1]
        
        if '--limit' in args:
            idx = args.index('--limit')
            if idx + 1 < len(args):
                limit = int(args[idx + 1])
        
        if '--start' in args:
            idx = args.index('--start')
            if idx + 1 < len(args):
                start_date = args[idx + 1]
        
        if '--end' in args:
            idx = args.index('--end')
            if idx + 1 < len(args):
                end_date = args[idx + 1]
        
        if not query:
            print("❌ Ошибка: Укажите --query")
            sys.exit(1)
        
        search_calendar(account, query, limit=limit, start_date=start_date, end_date=end_date)
    
    else:
        print(f"❌ Неизвестная команда: {command}")
        sys.exit(1)

if __name__ == "__main__":
    main()

