import os
import win32com.client
import openpyxl
import xlrd
import socket
from docx import Document
from odf.opendocument import load
from odf.text import P
from odf.table import Table, TableRow, TableCell


VERSION = '1.0.0'


def clean_text(text):
    return " ".join(text.split()).lower()


def search_in_tmp(file_path, search_texts):
    found_texts = []
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            text = clean_text(f.read())

            for phrase in search_texts:
                if clean_text(phrase) in text:
                    print(f"✅ Знайдено текст '{phrase}'")
                    found_texts.append(f"Знайдено текст '{phrase}'")
    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")
    
    return found_texts


def search_in_xlsx(file_path, search_texts):
    found_texts = []
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value:
                        cell_text = clean_text(str(cell.value))
                        for phrase in search_texts:
                            if clean_text(phrase) in cell_text:
                                print(f"✅ Знайдено текст '{phrase}'")
                                found_texts.append(f"Знайдено текст '{phrase}' (лист: {sheet.title})")
    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")
    return found_texts


def search_in_xls(file_path, search_texts):
    found_texts = []
    try:
        wb = xlrd.open_workbook(file_path)
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell(row_idx, col_idx).value
                    if isinstance(cell_value, str):
                        cell_text = clean_text(cell_value)
                        for phrase in search_texts:
                            if clean_text(phrase) in cell_text:
                                print(f"✅ Знайдено текст '{phrase}'")
                                found_texts.append(f"Знайдено текст '{phrase}' (лист: {sheet.name})")
    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")
    return found_texts


def search_in_doc(file_path, search_texts):
    found_texts = []
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(os.path.abspath(file_path))
        text = doc.Content.Text

        for phrase in search_texts:
            if phrase.lower() in text.lower():
                print(f"✅ Знайдено текст '{phrase}'")
                found_texts.append(f'Знайдено текст "{phrase}"')

        doc.Close(False)
        word.Quit()
    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")
        try:
            word.Quit()
        except:
            pass
    return found_texts


def search_in_docx(file_path, search_texts):
    found_texts = []
    try:
        doc = Document(file_path)

        for para in doc.paragraphs:
            for phrase in search_texts:
                if phrase.lower() in para.text.lower():
                    print(f"✅ Знайдено текст '{phrase}'")
                    found_texts.append(f"Знайдено текст '{phrase}': {para.text}")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for phrase in search_texts:
                        if phrase.lower() in cell.text.lower():
                            print(f"✅ Знайдено текст '{phrase}'")
                            found_texts.append(f"Знайдено текст '{phrase}' в таблиці: {cell.text}")

    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")
    return found_texts


def extract_text_from_odt(doc):
    TEXT_NODE = 3
    text = []
    for p in doc.getElementsByType(P):
        text.append("".join(node.data for node in p.childNodes if node.nodeType == TEXT_NODE))
    return " ".join(text)


def search_in_odt(file_path, search_texts):
    found_texts = []
    try:
        doc = load(file_path)
        text = clean_text(extract_text_from_odt(doc))

        for phrase in search_texts:
            if clean_text(phrase) in text:
                print(f"✅ Знайдено текст '{phrase}'")
                found_texts.append(f"Знайдено текст '{phrase}'")
    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")

    return found_texts


def extract_text_from_ods(doc):
    TEXT_NODE = 3
    text = []
    for table in doc.getElementsByType(Table):
        for row in table.getElementsByType(TableRow):
            for cell in row.getElementsByType(TableCell):
                paragraphs = cell.getElementsByType(P)
                for p in paragraphs:
                    cell_text = "".join(node.data for node in p.childNodes if node.nodeType == TEXT_NODE)
                    text.append(cell_text)
    return " ".join(text)


def search_in_ods(file_path, search_texts):
    found_texts = []
    try:
        doc = load(file_path)
        text = clean_text(extract_text_from_ods(doc))

        for phrase in search_texts:
            if clean_text(phrase) in text:
                print(f"✅ Знайдено текст '{phrase}'")
                found_texts.append(f"Знайдено текст '{phrase}'")
    except Exception as e:
        print(f"❌ Помилка обробки {file_path}: {e}")

    return found_texts


def search_in_all_files(root_folder, search_texts):
    results = []

    for foldername, subfolders, filenames in os.walk(root_folder):
        for filename in filenames:
            file_path = os.path.join(foldername, filename)

            if filename.lower().endswith(".docx"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_docx(file_path, search_texts)
            elif filename.lower().endswith(".doc"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_doc(file_path, search_texts)
            elif filename.lower().endswith(".xlsx"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_xlsx(file_path, search_texts)
            elif filename.lower().endswith(".xls"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_xls(file_path, search_texts)
            elif filename.lower().endswith(".odt"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_odt(file_path, search_texts)
            elif filename.lower().endswith(".ods"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_ods(file_path, search_texts)
            elif filename.lower().endswith(".tmp"):
                print(f"🔍 Перевіряю: {file_path}")
                found_texts = search_in_tmp(file_path, search_texts)
            else:
                continue

            if found_texts:
                results.append((file_path, found_texts))

    return results


def get_computer_name():
    computer_name = socket.gethostname()
    return computer_name


if __name__ == "__main__":
    print(f"document_scanner v{VERSION} підтримує формати doc, docx, xls, xlsx, odt, ods\n")

    disks = input("Введіть букву диска для пошуку (наприклад, CDE): ").strip().upper()
    search_phrases = ["Для службового користування", "Таємно"]

    all_results = {}
    computer_name = get_computer_name()

    for disk in disks:
        root_folder = f"{disk}:\\"

        if os.path.exists(root_folder):
            print(f"\n🔍 Починаємо пошук на диску {root_folder} ...")
            results = search_in_all_files(root_folder, search_phrases)

            all_results[disk] = results  

            result_file = f"{computer_name}_{disk}.txt"
            with open(result_file, "w", encoding="utf-8") as f:
                for file, texts in results:
                    f.write(f"\n📄 Файл: {file}\n")
                    for text in texts:
                        f.write(f"🔹 {text}\n")

            print(f"✅ Знайдено файлів на диску {disk}: {len(results)}. Результати збережено в {result_file}")

        else:
            print(f"❌ Помилка: диск {disk}:\\ не знайдено!")

    input("\n🔹 Натисніть Enter, для завершення роботи програми...")
