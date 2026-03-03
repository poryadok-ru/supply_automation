import pandas as pd
import smtplib
from email.message import EmailMessage
from datetime import datetime
import os
import time
import glob
from python_calamine.pandas import pandas_monkeypatch


pandas_monkeypatch()

pd.options.mode.chained_assignment = None

os.chdir(os.path.dirname(os.path.abspath(__file__)))


def send_letter(file_path, add_message):
    if not file_path.endswith(os.sep):
        file_path = file_path + os.sep

    # ================= НАСТРОЙКИ =================
    SMTP_SERVER = "server"  # SMTP сервер Zimbra
    SMTP_PORT = 587                     # Обычно 587 (TLS) или 465 (SSL)
    EMAIL_ADDRESS = "email"
    EMAIL_PASSWORD = "pass"

    EXCEL_FILE = file_path +  "recipients.xlsx"
    ATTACHMENT_PATH = file_path +  "prices.xlsx"  # файл для вложения (если не нужен — удалите блок вложения)

    DELAY_BETWEEN_EMAILS = 2             # задержка между письмами (сек)

    # ✅ Месяцы вручную (без locale!)
    months = {
        1: "Январь",
        2: "Февраль",
        3: "Март",
        4: "Апрель",
        5: "Май",
        6: "Июнь",
        7: "Июль",
        8: "Август",
        9: "Сентябрь",
        10: "Октябрь",
        11: "Ноябрь",
        12: "Декабрь"
    }

    current_month = months[datetime.now().month]

    # ===== Чтение Excel =====

    df = pd.read_excel(EXCEL_FILE)

    if "Получатель" not in df.columns:
        raise ValueError("В Excel нет столбца 'Получатель'")

    recipients = df["Получатель"].dropna().tolist()

    print(f"Найдено получателей: {len(recipients)}")

    # ===== Подключение (как было в рабочем коде) =====

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()  # без ssl.create_default_context()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        for recipient in recipients:
            msg = EmailMessage()
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = recipient
            msg["Subject"] = f"Минимальные розничные цены для оптовых клиентов на {current_month} 2026 года"

            body = f"""Уважаемые партнёры!

Информируем вас о минимальных розничных ценах на {current_month} 2026 года, которые необходимо соблюдать при реализации товаров из прилагаемого списка.
Данные цены установлены с целью поддержания единой ценовой политики на рынке и защиты интересов всех участников дистрибуционной сети.

Обязательные условия:
При реализации товаров из списка во вложении розничная цена не может быть ниже указанной в таблице.
Действие цен распространяется на весь {current_month} 2026 года.

Важно:
Нарушение указанных минимальных цен недопустимо и может привести к пересмотру условий сотрудничества.

Благодарим за понимание!

С уважением,
Отдел ценообразования
Группа компаний «Спектр», г. Воронеж, ул. Волгоградская, 32
тел.: +7 (473) 233-00-00
www.sct.ru
            """

            msg.set_content(body)

            # Вложение
            if ATTACHMENT_PATH and os.path.exists(ATTACHMENT_PATH):
                with open(ATTACHMENT_PATH, "rb") as f:
                    msg.add_attachment(
                        f.read(),
                        maintype="application",
                        subtype="octet-stream",
                        filename=os.path.basename(ATTACHMENT_PATH)
                    )

            try:
                server.send_message(msg)
                print(f"✅ Отправлено: {recipient}")
            except Exception as e:
                print(f"❌ Ошибка отправки {recipient}: {e}")

            time.sleep(DELAY_BETWEEN_EMAILS)

    add_message('📨 Рассылка завершена.')
    



