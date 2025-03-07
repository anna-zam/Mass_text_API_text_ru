import requests
import pandas as pd
import time
from docx import Document


def read_texts_from_docx(file_path):
    """
    Считывает тексты из Word-документа.
    Каждая статья разделяется пустой строкой.
    """
    doc = Document(file_path)
    articles = []
    current_text = ""

    for paragraph in doc.paragraphs:
        # Пустая строка отделяет одну статью от другой
        if paragraph.text.strip() == "":
            if current_text.strip():
                articles.append(current_text.strip())
                current_text = ""
        else:
            current_text += paragraph.text + "\n"

    # Добавляем последний текст, если он не пуст
    if current_text.strip():
        articles.append(current_text.strip())

    return articles


def check_texts_bulk_textru(api_key, texts, output_path="text_ru_results.xlsx"):
    """
    Массовая проверка списка текстов через API text.ru.
    Результаты сохраняются в Excel.

    :param api_key: Ваш API-ключ text.ru.
    :param texts: Список строк (текстов) для проверки.
    :param output_path: Путь и название Excel-файла с результатами.
    """

    # Эндпоинт для отправки текста на проверку и для опроса статуса
    url = "https://api.text.ru/post"

    results = []
    for idx, text in enumerate(texts):
        print(f"Проверяется текст {idx+1}/{len(texts)}...")

        # Если нужно пропускать короткие тексты — раскомментируйте или поменяйте условие.
        # if len(text) < 50:
        #     print(f"Текст {idx + 1} слишком короткий. Пропускаем.")
        #     results.append({"Текст": text[:50], "Ошибка": "Текст слишком короткий"})
        #     continue

        # Отправка текста на проверку
        data = {
            "userkey": api_key,
            "text": text,
            "visible": "text",       # Чтобы можно было смотреть результат (опционально)
            "jsonvisible": "detail", # Чтобы получить детали (включая SEO)
            "json": 1
        }

        try:
            response = requests.post(url, data=data)
            response.raise_for_status()
            start_check_json = response.json()
        except Exception as e:
            print(f"Ошибка при отправке текста {idx+1}: {str(e)}")
            results.append({
                "Текст": text[:50],
                "Уникальность": "",
                "Заспамленность": "",
                "Вода": "",
                "Ссылка": "",
                "Ошибка": f"Ошибка при отправке: {e}"
            })
            continue

        # Проверяем, вернулся ли UID (ID проверки)
        text_uid = start_check_json.get("text_uid")
        if not text_uid:
            # Если UID не вернулся, возможно ошибка в API или ограничение
            error_msg = start_check_json.get("error_desc") or "Не получен text_uid"
            print(f"Не удалось получить UID для текста {idx+1}: {error_msg}")
            results.append({
                "Текст": text[:50],
                "Уникальность": "",
                "Заспамленность": "",
                "Вода": "",
                "Ссылка": "",
                "Ошибка": error_msg
            })
            continue

        # Теперь опрашиваем статус проверки
        status_data = None
        status_params = {
            "userkey": api_key,
            "uid": text_uid,
            "jsonvisible": "detail",
            "json": 1
        }

        # Ждём не более 12 * 5 = 60 секунд
        for attempt in range(12):
            time.sleep(5)  # Пауза между опросами
            try:
                status_response = requests.post(url, data=status_params)
                status_response.raise_for_status()
                status_data = status_response.json()
            except Exception as e:
                status_data = {"error": f"Ошибка при опросе статуса: {str(e)}"}
                break

            # Если проверка завершилась (есть поле "text_unique")
            if "text_unique" in status_data:
                break

            # Если в ответе ошибка, например "Текст ещё проверяется" (error_code=181)
            error_code = status_data.get("error_code")
            if error_code != 181:
                # Значит произошла другая ошибка, выходим
                break

        # Сбор результатов
        if not status_data or "text_unique" not in status_data:
            # Проверка так и не завершилась или произошла ошибка
            results.append({
                "Текст": text[:50],
                "Уникальность": "",
                "Заспамленность": "",
                "Вода": "",
                "Ссылка": f"https://text.ru/antiplagiat/{text_uid}",
                "Ошибка": status_data.get("error", "Проверка не завершилась")
            })
        else:
            # Проверка завершена, парсим данные
            text_unique = status_data.get("text_unique", "Нет данных")
            # SEO-данные (заспамленность, вода) лежат в поле "seo_check" (JSON-строка)
            spam_percent = None
            water_percent = None

            seo_raw = status_data.get("seo_check")
            if seo_raw:
                # Обычно это строка вида '{"water_percent":30,"spam_percent":25,"mixed_words":...}'
                import json
                try:
                    seo_parsed = json.loads(seo_raw)
                    spam_percent = seo_parsed.get("spam_percent", "Нет данных")
                    water_percent = seo_parsed.get("water_percent", "Нет данных")
                except json.JSONDecodeError:
                    spam_percent = "Ошибка разбора SEO"
                    water_percent = "Ошибка разбора SEO"

            result_item = {
                "Текст": text[:50],
                "Уникальность": f"{text_unique}%",
                "Заспамленность": f"{spam_percent}%" if spam_percent is not None else "",
                "Вода": f"{water_percent}%" if water_percent is not None else "",
                "Ссылка": f"https://text.ru/antiplagiat/{text_uid}",
                "Ошибка": ""
            }
            results.append(result_item)

    # Сохраняем в Excel
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False)
    print(f"Результаты сохранены в файл: {output_path}")


# Пример использования
if __name__ == "__main__":
    # Замените API-ключ на ваш
    api_key = "Ваш_API_ключ"

    docx_path = "articles.docx"  # Путь к вашему Word-файлу
    texts = read_texts_from_docx(docx_path)

    check_texts_bulk_textru(api_key, texts, output_path="text_ru_results.xlsx")

