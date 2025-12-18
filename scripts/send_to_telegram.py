#!/usr/bin/env python3
"""
Скрипт для отправки файлов в Telegram из командной строки.

Требуется настройка:
1. Создать бота через @BotFather в Telegram
2. Получить токен бота
3. Получить Chat ID (можно через @userinfobot или отправить сообщение боту и проверить через API)
4. Сохранить токен и Chat ID в config.json (создается автоматически при первом запуске)
"""

import sys
import os
import json
from pathlib import Path

# Проверка наличия библиотеки requests
try:
    import requests
except ImportError:
    print("❌ Ошибка: Библиотека 'requests' не установлена.")
    print("   Установите её командой: pip3 install requests")
    sys.exit(1)

# Путь к конфигурационному файлу
CONFIG_FILE = Path(__file__).parent / "telegram_config.json"

def load_config():
    """Загружает конфигурацию из файла или создает шаблон."""
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            
            # Миграция со старого формата (если есть chat_id напрямую)
            if 'chat_id' in config and 'chats' not in config:
                print("⚠️  Обнаружен старый формат конфигурации. Мигрирую...")
                old_chat_id = config.pop('chat_id')
                config['chats'] = {
                    'default': old_chat_id if old_chat_id != "YOUR_CHAT_ID_HERE" else "YOUR_CHAT_ID_HERE"
                }
                config['default_chat'] = 'default'
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)
                print("✅ Конфигурация обновлена до нового формата")
            
            return config
    else:
        # Создаем шаблон конфигурации
        template = {
            "bot_token": "YOUR_BOT_TOKEN_HERE",
            "chats": {
                "myself": "YOUR_CHAT_ID_HERE",
                "doctor": "YOUR_CHAT_ID_HERE",
                "family": "YOUR_CHAT_ID_HERE"
            },
            "default_chat": "myself"
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(template, f, indent=2, ensure_ascii=False)
        print(f"⚠️  Создан файл конфигурации: {CONFIG_FILE}")
        print("📝 Пожалуйста, заполните bot_token и chat_id в файле конфигурации.")
        print("\nИнструкция:")
        print("1. Создайте бота через @BotFather в Telegram")
        print("2. Получите токен бота")
        print("3. Получите Chat ID для каждого чата (можно через @userinfobot)")
        print("4. Добавьте чаты в секцию 'chats' с понятными именами")
        print("5. Установите 'default_chat' на имя чата по умолчанию")
        return None

def get_chat_id(config, chat_name=None):
    """Получает Chat ID или username по имени или использует дефолтный."""
    if not chat_name:
        chat_name = config.get('default_chat', 'myself')
    
    chats = config.get('chats', {})
    if chat_name not in chats:
        print(f"❌ Ошибка: Чат '{chat_name}' не найден в конфигурации.")
        print(f"   Доступные чаты: {', '.join(chats.keys())}")
        return None
    
    chat_id = chats[chat_name]
    if chat_id == "YOUR_CHAT_ID_HERE":
        print(f"❌ Ошибка: Chat ID для '{chat_name}' не настроен.")
        return None
    
    # Поддержка username (начинается с @)
    if chat_id.startswith('@'):
        return chat_id
    
    return chat_id

def send_file_to_telegram(file_path, bot_token, chat_id, caption=None):
    """Отправляет файл в Telegram."""
    if not os.path.exists(file_path):
        print(f"❌ Ошибка: Файл не найден: {file_path}")
        return False
    
    url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
    
    try:
        with open(file_path, 'rb') as file:
            files = {'document': (os.path.basename(file_path), file)}
            data = {'chat_id': chat_id}
            if caption:
                data['caption'] = caption
            
            response = requests.post(url, files=files, data=data, timeout=30)
            response.raise_for_status()
            
            result = response.json()
            if result.get('ok'):
                print(f"✅ Файл успешно отправлен в Telegram!")
                print(f"   Файл: {os.path.basename(file_path)}")
                return True
            else:
                print(f"❌ Ошибка отправки: {result.get('description', 'Unknown error')}")
                return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при отправке: {e}")
        return False
    except Exception as e:
        print(f"❌ Неожиданная ошибка: {e}")
        return False

def send_text_to_telegram(text, bot_token, chat_id):
    """Отправляет текстовое сообщение в Telegram."""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    
    try:
        data = {
            'chat_id': chat_id,
            'text': text,
            'parse_mode': 'Markdown'
        }
        
        response = requests.post(url, data=data, timeout=30)
        
        # Пытаемся получить JSON ответ даже при ошибке
        try:
            result = response.json()
        except:
            result = {}
        
        if response.status_code == 200 and result.get('ok'):
            print(f"✅ Сообщение успешно отправлено в Telegram!")
            return True
        else:
            error_desc = result.get('description', response.text or 'Unknown error')
            error_code = result.get('error_code', '')
            print(f"❌ Ошибка отправки: {error_desc}")
            if error_code:
                print(f"   Код ошибки: {error_code}")
            # Полезные подсказки для частых ошибок
            if "chat not found" in error_desc.lower() or "chat_id" in error_desc.lower():
                print("   💡 Подсказка: Убедитесь, что вы отправили /start боту")
                print("      Или проверьте правильность Chat ID")
            elif "unauthorized" in error_desc.lower():
                print("   💡 Подсказка: Проверьте правильность токена бота")
            return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка при отправке: {e}")
        return False
    except Exception as e:
        print(f"❌ Неожиданная ошибка: {e}")
        return False

def list_chats(config):
    """Выводит список доступных чатов."""
    chats = config.get('chats', {})
    default = config.get('default_chat', 'myself')
    
    print("📋 Доступные чаты:")
    for name, chat_id in chats.items():
        marker = " (по умолчанию)" if name == default else ""
        if chat_id == "YOUR_CHAT_ID_HERE":
            status = "❌ не настроен"
        elif chat_id.startswith('@'):
            status = f"✅ настроен (username: {chat_id})"
        else:
            status = "✅ настроен (Chat ID)"
        print(f"   • {name}: {status}{marker}")

def main():
    # Парсим аргументы
    chat_name = None
    args = sys.argv[1:]
    
    # Обработка флагов
    if '--list' in args:
        config = load_config()
        if config:
            list_chats(config)
        sys.exit(0)
    
    if '--chat' in args:
        idx = args.index('--chat')
        if idx + 1 >= len(args):
            print("❌ Ошибка: Укажите имя чата после --chat")
            sys.exit(1)
        chat_name = args[idx + 1]
        args = [a for a in args if a != '--chat' and a != chat_name]
    
    if len(args) < 1:
        print("Использование:")
        print(f"  {sys.argv[0]} <путь_к_файлу> [подпись] [--chat <имя_чата>]")
        print(f"  {sys.argv[0]} --text '<текст>' [--chat <имя_чата>]")
        print(f"  {sys.argv[0]} --list  # показать список чатов")
        print("\nПримеры:")
        print(f"  {sys.argv[0]} /path/to/file.txt")
        print(f"  {sys.argv[0]} /path/to/file.txt 'Результаты анализов'")
        print(f"  {sys.argv[0]} /path/to/file.txt --chat doctor")
        print(f"  {sys.argv[0]} --text 'Привет!' --chat doctor")
        print(f"  {sys.argv[0]} --list")
        sys.exit(1)
    
    # Загружаем конфигурацию
    config = load_config()
    if not config:
        sys.exit(1)
    
    bot_token = config.get('bot_token')
    
    if bot_token == "YOUR_BOT_TOKEN_HERE":
        print("❌ Ошибка: Необходимо настроить bot_token в файле конфигурации.")
        print(f"   Файл: {CONFIG_FILE}")
        sys.exit(1)
    
    # Получаем Chat ID
    chat_id = get_chat_id(config, chat_name)
    if not chat_id:
        sys.exit(1)
    
    # Проверяем, отправляем ли мы текст
    if args[0] == '--text':
        if len(args) < 2:
            print("❌ Ошибка: Укажите текст сообщения после --text")
            sys.exit(1)
        text = args[1]
        chat_display = chat_name if chat_name else config.get('default_chat', 'default')
        print(f"📤 Отправка сообщения в чат: {chat_display}")
        send_text_to_telegram(text, bot_token, chat_id)
    else:
        # Отправляем файл
        file_path = args[0]
        caption = args[1] if len(args) > 1 and args[1] != '--chat' else None
        chat_display = chat_name if chat_name else config.get('default_chat', 'default')
        print(f"📤 Отправка файла в чат: {chat_display}")
        send_file_to_telegram(file_path, bot_token, chat_id, caption)

if __name__ == "__main__":
    main()

