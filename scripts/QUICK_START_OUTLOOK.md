# 🚀 Быстрый старт: Работа с Office 365 почтой

## Шаг 1: Установите зависимости

```bash
pip3 install O365
```

## Шаг 2: Зарегистрируйте приложение в Azure

1. Откройте [Azure Portal](https://portal.azure.com/)
2. Azure Active Directory → App registrations → New registration
3. Настройте приложение и скопируйте **Application (client) ID**
4. Certificates & secrets → New client secret → скопируйте **Value** (показывается один раз!)
5. API permissions → Add permission → Microsoft Graph → Delegated permissions:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
6. Grant admin consent (если требуется)

## Шаг 3: Настройте скрипт

```bash
# Запустите один раз - создастся файл конфигурации
python3 scripts/outlook_email.py list
```

Откройте `scripts/outlook_config.json` и заполните:
```json
{
  "client_id": "ваш_client_id_из_azure",
  "client_secret": "ваш_client_secret_из_azure",
  "tenant_id": "common",
  "scopes": ["basic", "message_all"],
  "email": "ваш_email@domain.com"
}
```

## Шаг 4: Первая аутентификация

При первом запуске откроется браузер:
1. Войдите в Office 365
2. Предоставьте разрешения
3. Токен сохранится автоматически

## Шаг 5: Используйте!

```bash
# Просмотр писем
python3 scripts/outlook_email.py list --limit 10

# Чтение письма
python3 scripts/outlook_email.py read --index 0

# Отправка письма
python3 scripts/outlook_email.py send \
  --to "recipient@example.com" \
  --subject "Тема" \
  --body "Текст письма"

# Поиск
python3 scripts/outlook_email.py search --query "важно"
```

**Готово!** 🎉

---

📖 **Подробная инструкция:** `scripts/outlook_setup.md`

