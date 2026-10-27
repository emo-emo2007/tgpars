# Telegram-посты в Word

Скрипт загружает текст и фотографии из публичных постов Telegram и сохраняет их в DOCX: ссылка, фотографии, текст. Шрифт текста документа — Times New Roman, 12 pt.

## Установка

```powershell
python -m pip install -r requirements.txt
```

## Запуск

Создайте `links.txt` и поместите в него по одной ссылке на пост в каждой строке. Пустые строки и строки, начинающиеся с `#`, игнорируются.

```text
https://t.me/tr_sputnik/175574
```

Затем запустите:

```powershell
python telegram_to_word.py links.txt -o posts.docx
```

Если пост виден только после входа в Telegram, получите API ID и API hash на [my.telegram.org](https://my.telegram.org), затем задайте их в PowerShell и запустите авторизованный режим:

```powershell
$env:TELEGRAM_API_ID = "ваш_api_id"
$env:TELEGRAM_API_HASH = "ваш_api_hash"
python telegram_to_word.py links.txt -o posts.docx --telegram-auth
```

При первом запуске Telegram запросит номер телефона, код подтверждения и, если включена, пароль двухфакторной аутентификации. Код и пароль вводятся только в терминале. Файл сессии сохраняется в домашней папке пользователя, не в папке проекта. Удалённые посты и посты, к которым у аккаунта нет доступа, получить нельзя. Если часть ссылок недоступна, остальные всё равно попадут в документ.