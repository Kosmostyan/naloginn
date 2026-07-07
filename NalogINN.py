import pandas as pd
import requests
import time
import os
from datetime import datetime
import re


def format_document_number(docno_full):
    cleaned = re.sub(r'[\s\-_/]', '', str(docno_full))
    cleaned = re.sub(r'\D', '', cleaned)

    if len(cleaned) < 10:
        print(f"Номер документа слишком короткий: {cleaned} (длина {len(cleaned)})")
        return None

    cleaned = cleaned[:10]
    series = cleaned[:4]
    number = cleaned[4:10]
    formatted = f"{series[:2]} {series[2:4]} {number}"
    return formatted


def get_inn_by_person_data(fam, nam, otch, bdate, doctype_code, docno_full, docdt):
    session = requests.Session()

    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Accept-Language': 'ru,en;q=0.9',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Host': 'service.nalog.ru',
        'Origin': 'https://service.nalog.ru',
        'Referer': 'https://service.nalog.ru/inn.do',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 YaBrowser/26.3.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
    }

    url_inn_proc = "https://service.nalog.ru/inn-new-proc.json"

    formatted_docno = format_document_number(docno_full)
    if not formatted_docno:
        print(f"Ошибка: Не удалось отформатировать номер документа '{docno_full}'")
        return None

    print(f"Исходный номер: {docno_full} - Отформатированный: {formatted_docno}")

    payload1 = {
        "c": "find",
        "fam": fam,
        "nam": nam,
        "otch": otch,
        "bdate": bdate,
        "doctype": doctype_code,
        "docno": formatted_docno,
        "docdt": docdt,
        "captcha": "",
        "captchaToken": ""
    }

    try:
        print(f"Отправляем первый запрос для: {fam} {nam} {otch}")
        response1 = session.post(url_inn_proc, data=payload1, headers=headers)

        print(f"Статус первого запроса: {response1.status_code}")

        if response1.status_code != 200:
            print(f"Ошибка первого запроса: {response1.text}")
            return None

        json_response1 = response1.json()

        if not json_response1 or "requestId" not in json_response1 or "token1" not in json_response1:
            print(f"Ошибка: Не удалось получить requestId или token1")
            return None

        request_id = json_response1["requestId"]
        token1 = json_response1["token1"]
        print(f"Получен requestId: {request_id[:50]}...")
        print(f"Получен token1: {token1[:50]}...")

        print("Ожидание 10 секунд перед проверкой результата...")
        time.sleep(10)

        payload2 = {
            "c": "get",
            "requestId": request_id,
            "token": token1
        }

        print("Отправляем второй запрос для получения результата...")

        response2 = session.post(url_inn_proc, data=payload2, headers=headers)

        print(f"Статус второго запроса: {response2.status_code}")

        if response2.status_code == 200:
            try:
                json_response2 = response2.json()

                if "inn" in json_response2 and json_response2["inn"]:
                    found_inn = json_response2["inn"]
                    print(f"Найден ИНН: {found_inn}")
                    return found_inn
                else:
                    print(f"ИНН не найден в ответе")
                    return None

            except Exception as e:
                print(f"Ошибка парсинга JSON: {e}")
                return None
        else:
            print(f"Ошибка второго запроса: {response2.status_code}")
            return None

    except Exception as e:
        print(f"Непредвиденная ошибка: {e}")
        return None


def parse_date(date_value):
    try:
        if isinstance(date_value, datetime):
            return date_value.strftime('%d.%m.%Y')

        if isinstance(date_value, str):
            date_str = date_value.strip()

            formats = [
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d',
                '%d.%m.%Y',
                '%d/%m/%Y',
                '%d-%m-%Y'
            ]

            for fmt in formats:
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    return date_obj.strftime('%d.%m.%Y')
                except ValueError:
                    continue

        try:
            date_obj = pd.to_datetime(date_value)
            return date_obj.strftime('%d.%m.%Y')
        except:
            pass

        return None

    except Exception as e:
        print(f"Ошибка парсинга даты '{date_value}': {e}")
        return None


def main():
    script_directory = os.path.dirname(os.path.abspath(__file__))
    excel_files = [f for f in os.listdir(script_directory) if f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print(f"Ошибка: В той же папке, что и скрипт, не найдено ни одного Excel-файла.")
        return

    original_excel_file = os.path.join(script_directory, excel_files[0])
    base, ext = os.path.splitext(original_excel_file)
    new_excel_file = f"{base}_processed{ext}"

    print(f"Найден исходный Excel-файл: {original_excel_file}")
    print(f"Новый Excel-файл будет создан: {new_excel_file}")

    INPUT_COLUMNS = {
        'fam': 'Фамилия',
        'nam': 'Имя',
        'otch': 'Отчество',
        'bdate': 'Дата рождения',
        'doctype_text': 'Вид документа, удостоверяющего личность',
        'docno_full': 'Серия и номер документа',
        'docdt': 'Дата выдачи документа'
    }
    OUTPUT_COLUMN_INN = 'ИНН'

    try:
        df = pd.read_excel(original_excel_file)

        for key, col_name in INPUT_COLUMNS.items():
            if col_name not in df.columns:
                print(f"Критическая ошибка: В файле отсутствует колонка '{col_name}'.")
                return

        for col_name in INPUT_COLUMNS.values():
            df[col_name] = df[col_name].astype(str)

        if OUTPUT_COLUMN_INN not in df.columns:
            df[OUTPUT_COLUMN_INN] = ''
        df[OUTPUT_COLUMN_INN] = df[OUTPUT_COLUMN_INN].astype(object)

    except FileNotFoundError:
        print(f"Ошибка: Файл '{original_excel_file}' не найден.")
        return
    except Exception as e:
        print(f"Ошибка при чтении Excel файла: {e}")
        return

    print(f"Начинаем обработку строк...")

    total_rows = len(df)
    processed_count = 0

    for index, row in df.iterrows():
        print(f"\nОбработка строки {index + 1}/{total_rows}")

        fam = row.get(INPUT_COLUMNS['fam'], '').strip()
        nam = row.get(INPUT_COLUMNS['nam'], '').strip()
        otch = row.get(INPUT_COLUMNS['otch'], '').strip()
        bdate = row.get(INPUT_COLUMNS['bdate'], '').strip()
        doctype_text = row.get(INPUT_COLUMNS['doctype_text'], '').strip()
        docno_full = row.get(INPUT_COLUMNS['docno_full'], '').strip()
        docdt = row.get(INPUT_COLUMNS['docdt'], '').strip()

        if (fam in ['', 'nan', 'None'] or nam in ['', 'nan', 'None'] or
                bdate in ['', 'nan', 'None'] or docno_full in ['', 'nan', 'None']):
            print(f"Пропускаем из-за неполных данных.")
            df.at[index, OUTPUT_COLUMN_INN] = 'Неполные данные'
            try:
                df.to_excel(new_excel_file, index=False)
            except Exception as e:
                print(f"Ошибка сохранения: {e}")
            continue

        doctype_code = '21'
        if doctype_text and doctype_text != 'nan':
            match = re.search(r'(\d+)', doctype_text)
            if match:
                doctype_code = match.group(1)

        print(f"Код документа: {doctype_code}")

        bdate_formatted = parse_date(bdate)
        if not bdate_formatted:
            print(f"Не удалось распознать дату рождения: {bdate}")
            df.at[index, OUTPUT_COLUMN_INN] = 'Ошибка даты рождения'
            continue

        docdt_formatted = parse_date(docdt)
        if not docdt_formatted:
            print(f"Не удалось распознать дату выдачи: {docdt}")
            df.at[index, OUTPUT_COLUMN_INN] = 'Ошибка даты выдачи'
            continue

        print(f"Дата рождения: {bdate_formatted}")
        print(f"Дата выдачи: {docdt_formatted}")

        found_inn = get_inn_by_person_data(
            fam=fam,
            nam=nam,
            otch=otch,
            bdate=bdate_formatted,
            doctype_code=doctype_code,
            docno_full=docno_full,
            docdt=docdt_formatted
        )

        if found_inn:
            df.at[index, OUTPUT_COLUMN_INN] = found_inn
        else:
            df.at[index, OUTPUT_COLUMN_INN] = 'ИНН не найден'

        processed_count += 1
        print(f"Обработано строк: {processed_count}/{total_rows}")

        if index < total_rows - 1:
            print("Ожидание 10 секунд перед следующим запросом...")
            time.sleep(10)

        try:
            df.to_excel(new_excel_file, index=False)
            print(f"Промежуточное сохранение выполнено")
        except Exception as e:
            print(f"Ошибка сохранения: {e}")

    print(f"\nОбработка завершена. Обработано строк: {processed_count}/{total_rows}")
    print(f"Результат сохранен в: {new_excel_file}")


if __name__ == "__main__":
    main()