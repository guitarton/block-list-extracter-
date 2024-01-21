import imaplib
import email
import re
from email.header import decode_header
from datetime import datetime
from openpyxl import Workbook
from datetime import datetime
from urlextract import URLExtract

# Параметры для подключения к почтовому серверу
imap_server = "192.168.199.39"
username = "bob@mail.local"
password = "Pa$$w0rd"

subject_pattern = re.compile(r"Добавление индикаторов на блокировку", re.IGNORECASE)

# Создание объекта для подключения к почтовому серверу
mail = imaplib.IMAP4_SSL(imap_server)

# Аутентификация
mail.login(username, password)

# Выбор папки с непрочитанными письмами
mail.select("inbox")
status, messages = mail.search(None, "(UNSEEN)")

# Создание Excel-файла
timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")[:-3]
excel_file_path = f"result-{timestamp}.xlsx"
workbook = Workbook()
sheet = workbook.active
sheet.append(["TimeSent", "Sender", "Subject", "Body", "URLs"])


# Обработка каждого непрочитанного письма
for num in messages[0].split():
    status, msg_data = mail.fetch(num, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])

            # Проверка темы письма с использованием регулярного выражения
            subject, encoding = decode_header(msg.get("Subject", ""))[0]
            subject = subject.decode(encoding) if encoding else subject
            if not subject_pattern.search(subject):
                continue

            time_sent = datetime.strptime(msg["Date"], "%a, %d %b %Y %H:%M:%S %z")
            sender, encoding = decode_header(msg.get("From", ""))[0]
            sender = sender.decode(encoding) if encoding else sender
            body = ""

            # Получение текстового содержимого письма
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8")
            else:
                body = msg.get_payload(decode=True).decode("utf-8")

            # Извлечение IP-адресов и URL-адресов
            extractor = URLExtract()
            urls = ", ".join(extractor.find_urls(body))

            # Запись данных в Excel
            sheet.append([str(time_sent), sender, subject, body, urls])

# Сохранение Excel-файла
workbook.save(excel_file_path)

# Закрытие соединения с почтовым сервером
mail.logout()

print("Скрипт успешно выполнен.")
