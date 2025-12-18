# Настройка работы с Office 365 через Exchange Web Services (EWS)

Этот вариант использует протокол Exchange и **не требует Azure регистрации**. Работает напрямую с Exchange сервером Office 365.

## Шаг 1: Установка библиотеки

```bash
pip3 install exchangelib
```

## Шаг 2: Создание пароля приложения (если включена 2FA)

Если у вас включена двухфакторная аутентификация:

1. Перейдите на [account.microsoft.com/security](https://account.microsoft.com/security)
2. Войдите в свой аккаунт Microsoft
3. Выберите **Безопасность** → **Дополнительные параметры безопасности**
4. Найдите **Пароли приложений** (App passwords)
5. Создайте новый пароль приложения
6. **Скопируйте пароль** — он показывается только один раз!

**Важно:** Используйте этот пароль приложения вместо обычного пароля в конфигурации.

## Шаг 3: Настройка конфигурации

1. Запустите скрипт один раз (он создаст файл конфигурации):
   ```bash
   python3 scripts/outlook_exchange.py list
   ```

2. Откройте файл `scripts/outlook_exchange_config.json`

3. Заполните данные:
   ```json
   {
     "email": "ваш_email@domain.com",
     "username": "DOMAIN\\username",
     "password": "ваш_пароль_или_пароль_приложения",
     "server": "outlook.office365.com",
     "port": 443,
     "autodiscover": true,
     "verify_ssl": true
   }
   ```

   **Параметры:**
   - `email` - ваш email адрес
   - `username` - (опционально) логин в формате `DOMAIN\username` для корпоративных серверов
   - `password` - пароль от почты (или пароль приложения, если включена 2FA)
   - `server` - адрес Exchange сервера
   - `port` - (опционально) порт сервера (по умолчанию 443)
   - `autodiscover` - использовать автоматическое определение сервера (true/false)
   - `verify_ssl` - проверять SSL сертификат (true/false)

   **Пример для корпоративного сервера:**
   ```json
   {
     "email": "g.panibratets@alataucitybank.kz",
     "username": "tsb\\u00023788",
     "password": "ваш_пароль",
     "server": "mail.jusan.kz",
     "port": 443,
     "autodiscover": false,
     "verify_ssl": false
   }
   ```

   **Важно:** 
   - Для корпоративных серверов обычно требуется указать `username` в формате `DOMAIN\username`
   - В JSON обратный слэш нужно экранировать: `"tsb\\u00023788"` (два обратных слэша)
   - Если сервер использует самоподписанный сертификат, установите `verify_ssl: false`

## Шаг 4: Использование

### Просмотр списка писем:
```bash
# Последние 10 писем
python3 scripts/outlook_exchange.py list

# Последние 5 писем
python3 scripts/outlook_exchange.py list --limit 5

# Письма из отправленных
python3 scripts/outlook_exchange.py list --folder sent

# Письма из черновиков
python3 scripts/outlook_exchange.py list --folder drafts
```

### Чтение письма:
```bash
# Прочитать первое письмо
python3 scripts/outlook_exchange.py read --index 0

# Прочитать второе письмо
python3 scripts/outlook_exchange.py read --index 1
```

### Отправка письма:
```bash
# Простое письмо
python3 scripts/outlook_exchange.py send \
  --to "recipient@example.com" \
  --subject "Тема письма" \
  --body "Текст письма"

# Письмо с вложением
python3 scripts/outlook_exchange.py send \
  --to "recipient@example.com" \
  --subject "Документ" \
  --body "См. вложение" \
  --attach "/path/to/file.pdf"

# Письмо с копией
python3 scripts/outlook_exchange.py send \
  --to "recipient@example.com" \
  --subject "Тема" \
  --body "Текст" \
  --cc "copy@example.com"
```

### Поиск писем:
```bash
# Поиск по запросу
python3 scripts/outlook_exchange.py search --query "важно"

# Поиск с ограничением
python3 scripts/outlook_exchange.py search --query "проект" --limit 5
```

## Преимущества этого метода:

✅ Не требует Azure Portal  
✅ Не требует административных разрешений  
✅ Использует нативный протокол Exchange  
✅ Работает с корпоративными аккаунтами  
✅ Автоматическое определение сервера (autodiscover)  

## Безопасность:

- Файл `outlook_exchange_config.json` автоматически добавлен в `.gitignore`
- Пароль не будет закоммичен в Git
- Не делитесь паролем с другими

## Устранение проблем

### Ошибка: "Invalid credentials"
Для корпоративных Exchange серверов может потребоваться:

1. **Доменная аутентификация:**
   ```json
   {
     "email": "username@domain.com",
     "password": "your_password",
     "domain": "DOMAIN_NAME",
     "server": "mail.domain.com",
     "autodiscover": false
   }
   ```

2. **Полный email как username:**
   - Убедитесь, что используете полный email адрес
   - Проверьте правильность пароля

3. **Проверьте настройки Outlook:**
   - Посмотрите, какой формат учетных данных использует Outlook на вашем Mac
   - Возможно, требуется формат `DOMAIN\username` или другой

### Ошибка: "ErrorInvalidUser"
- Проверьте правильность email и пароля
- Если включена 2FA, используйте пароль приложения
- Убедитесь, что Exchange доступен для вашего аккаунта
- Для корпоративных серверов может потребоваться доменное имя

### Ошибка: "Autodiscover failed"
- Попробуйте указать сервер вручную в конфигурации:
  ```json
  {
    "autodiscover": false,
    "server": "mail.domain.com"
  }
  ```

### Ошибка: "SSL certificate verification failed"
- Установите `"verify_ssl": false` в конфигурации (для корпоративных серверов)
- Скрипт автоматически обрабатывает SSL, но если проблемы сохраняются, проверьте настройки сети/прокси

