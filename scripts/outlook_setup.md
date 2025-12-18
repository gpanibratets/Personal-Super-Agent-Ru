# Настройка работы с Office 365 почтой

Этот скрипт позволяет читать и отправлять письма из Office 365/Outlook прямо из командной строки через Microsoft Graph API.

## Шаг 1: Регистрация приложения в Azure Portal

1. Откройте [Azure Portal](https://portal.azure.com/)
2. Войдите с учетными данными Office 365
3. Перейдите в **Azure Active Directory** (или **Microsoft Entra ID**)
4. Выберите **App registrations** (Регистрация приложений)
5. Нажмите **New registration** (Новая регистрация)

### Настройка приложения:

- **Name** (Имя): `Personal Email Script` (или любое другое)
- **Supported account types** (Поддерживаемые типы учетных записей):
  - Выберите "Accounts in any organizational directory and personal Microsoft accounts"
- **Redirect URI** (URI перенаправления):
  - Platform: **Public client/native (mobile & desktop)**
  - URI: `http://localhost:8080` (или любой другой локальный адрес)

6. Нажмите **Register** (Зарегистрировать)

## Шаг 2: Получение Client ID и Client Secret

### Client ID (Application ID):
1. После регистрации вы увидите страницу приложения
2. Скопируйте **Application (client) ID** — это ваш Client ID

### Client Secret:
1. В меню приложения выберите **Certificates & secrets** (Сертификаты и секреты)
2. Нажмите **New client secret** (Новый секрет клиента)
3. Введите описание (например: "Email Script Secret")
4. Выберите срок действия (Expires)
5. Нажмите **Add** (Добавить)
6. **ВАЖНО:** Скопируйте **Value** секрета сразу — он показывается только один раз!

## Шаг 3: Настройка разрешений (API Permissions)

1. В меню приложения выберите **API permissions** (Разрешения API)
2. Нажмите **Add a permission** (Добавить разрешение)
3. Выберите **Microsoft Graph**
4. Выберите **Delegated permissions** (Делегированные разрешения)
5. Добавьте следующие разрешения:
   - `Mail.Read` — чтение почты
   - `Mail.ReadWrite` — чтение и запись почты
   - `Mail.Send` — отправка почты
   - `User.Read` — чтение профиля пользователя
6. Нажмите **Add permissions** (Добавить разрешения)
7. Нажмите **Grant admin consent** (Предоставить согласие администратора), если требуется

## Шаг 4: Настройка конфигурации

1. Запустите скрипт один раз (он создаст файл конфигурации):
   ```bash
   python3 scripts/outlook_email.py list
   ```

2. Откройте файл `scripts/outlook_config.json`

3. Заполните данные:
   ```json
   {
     "client_id": "ваш_client_id_из_azure",
     "client_secret": "ваш_client_secret_из_azure",
     "tenant_id": "common",
     "scopes": ["basic", "message_all"],
     "email": "ваш_email@domain.com"
   }
   ```

## Шаг 5: Первая аутентификация

При первом запуске скрипт откроет браузер для аутентификации:

1. Войдите в свой аккаунт Office 365
2. Предоставьте разрешения приложению
3. После успешной аутентификации токен будет сохранен локально
4. В дальнейшем аутентификация будет автоматической

## Шаг 6: Использование

### Просмотр списка писем:
```bash
# Последние 10 писем из входящих
python3 scripts/outlook_email.py list

# Последние 5 писем
python3 scripts/outlook_email.py list --limit 5

# Письма из отправленных
python3 scripts/outlook_email.py list --folder sent
```

### Чтение письма:
```bash
# Прочитать первое письмо (индекс 0)
python3 scripts/outlook_email.py read --index 0

# Прочитать второе письмо
python3 scripts/outlook_email.py read --index 1

# Прочитать письмо по ID
python3 scripts/outlook_email.py read --id MESSAGE_ID
```

### Отправка письма:
```bash
# Простое письмо
python3 scripts/outlook_email.py send \
  --to "recipient@example.com" \
  --subject "Тема письма" \
  --body "Текст письма"

# Письмо с вложением
python3 scripts/outlook_email.py send \
  --to "recipient@example.com" \
  --subject "Документ" \
  --body "См. вложение" \
  --attach "/path/to/file.pdf"

# Письмо с копией
python3 scripts/outlook_email.py send \
  --to "recipient@example.com" \
  --subject "Тема" \
  --body "Текст" \
  --cc "copy@example.com"
```

### Поиск писем:
```bash
# Поиск по запросу
python3 scripts/outlook_email.py search --query "важно"

# Поиск с ограничением
python3 scripts/outlook_email.py search --query "проект" --limit 5
```

## Безопасность

⚠️ **Важно:**
- Файл `outlook_config.json` автоматически добавлен в `.gitignore`
- Client Secret не будет закоммичен в Git
- Не делитесь Client Secret с другими
- Если секрет скомпрометирован, создайте новый в Azure Portal

## Устранение проблем

### Ошибка: "Invalid client"
- Проверьте правильность Client ID и Client Secret
- Убедитесь, что секрет не истек

### Ошибка: "Insufficient privileges"
- Проверьте, что разрешения настроены в Azure Portal
- Убедитесь, что предоставлено согласие администратора

### Ошибка: "Authentication failed"
- Удалите сохраненный токен (обычно в `~/.O365_token.txt`)
- Попробуйте аутентифицироваться заново

### Ошибка: "Token expired"
- Токен автоматически обновляется при использовании
- Если проблема сохраняется, удалите токен и аутентифицируйтесь заново

## Дополнительные ресурсы

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/)
- [O365 Python Library](https://github.com/O365/python-o365)
- [Azure Portal](https://portal.azure.com/)

