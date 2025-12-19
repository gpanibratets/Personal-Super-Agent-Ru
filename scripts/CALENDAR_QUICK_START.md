# 📅 Быстрый старт: Работа с календарем Exchange

## Просмотр событий

```bash
# Просмотр ближайших событий (следующие 30 дней)
python3 scripts/outlook_exchange.py calendar

# Ограничить количество
python3 scripts/outlook_exchange.py calendar --limit 10

# Конкретный период
python3 scripts/outlook_exchange.py calendar --start "2025-12-01" --end "2025-12-31"
```

## Создание встречи

```bash
# Простая встреча
python3 scripts/outlook_exchange.py calendar-create \
  --subject "Встреча с Profitbase" \
  --start "2025-12-24 09:00" \
  --end "2025-12-24 10:00"

# Встреча с участниками
python3 scripts/outlook_exchange.py calendar-create \
  --subject "Встреча с Profitbase" \
  --start "2025-12-24 09:00" \
  --end "2025-12-24 10:00" \
  --attendees "nshirobokova@profitbase.ru,S.kaisarov@alataucitybank.kz" \
  --body "Обсуждение интеграции API для онлайн-ипотеки" \
  --location "Онлайн"
```

## Поиск событий

```bash
# Поиск по ключевому слову
python3 scripts/outlook_exchange.py calendar-search --query "Profitbase"

# Поиск с ограничением периода
python3 scripts/outlook_exchange.py calendar-search \
  --query "ипотека" \
  --start "2025-12-01" \
  --end "2025-12-31"
```

## Форматы дат

- `YYYY-MM-DD HH:MM` - например: `2025-12-24 09:00`
- `YYYY-MM-DD` - дата без времени
- `YYYY-MM-DDTHH:MM:SS` - ISO формат

## Примеры использования

### Создание встречи на основе email переписки

Если вы получили email с предложением встречи, можно быстро создать событие:

```bash
python3 scripts/outlook_exchange.py calendar-create \
  --subject "Встреча с Максимом Селезневым (Profitbase)" \
  --start "2025-12-24 09:00" \
  --end "2025-12-24 10:00" \
  --attendees "nshirobokova@profitbase.ru" \
  --body "Обсуждение интеграции API Profitbase с банком для онлайн-ипотеки"
```

### Поиск встреч на конкретную дату

```bash
# Найти все встречи 24 декабря
python3 scripts/outlook_exchange.py calendar \
  --start "2025-12-24" \
  --end "2025-12-24"
```

### Поиск встреч с конкретным человеком

```bash
# Поиск по теме (если имя в теме)
python3 scripts/outlook_exchange.py calendar-search --query "Надежда"
```

---

**Подробная документация:** См. `scripts/README.md`
