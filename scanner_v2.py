import os
import win32com.client
import openpyxl
import xlrd
from docx import Document


def clean_text(text):
    return " ".join(text.split()).lower()


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
                                print(f'✅ Знайдено текст "{phrase}"')
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
                                print(f'✅ Знайдено текст "{phrase}"')
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
                print(f'✅ Знайдено текст "{phrase}"')
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
                    print(f'✅ Знайдено текст "{phrase}"')
                    found_texts.append(f"Знайдено текст '{phrase}': {para.text}")

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for phrase in search_texts:
                        if phrase.lower() in cell.text.lower():
                            print(f'✅ Знайдено текст "{phrase}"')
                            found_texts.append(f"Знайдено текст '{phrase}' в таблиці: {cell.text}")

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
            else:
                continue

            if found_texts:
                results.append((file_path, found_texts))

    return results


if __name__ == "__main__":
    disks = input("Введіть букву диска для пошуку (наприклад, CDE): ").strip().upper()
    search_phrases = ["Для службового користування", "Таємно"]

    all_results = {}

    for disk in disks:
        root_folder = f"{disk}:\\"

        if os.path.exists(root_folder):
            print(f"\n🔍 Починаємо пошук на диску {root_folder} ...")
            results = search_in_all_files(root_folder, search_phrases)

            all_results[disk] = results  

            result_file = f"search_results_{disk}.txt"
            with open(result_file, "w", encoding="utf-8") as f:
                for file, texts in results:
                    f.write(f"\n📄 Файл: {file}\n")
                    for text in texts:
                        f.write(f"🔹 {text}\n")

            print(f"✅ Знайдено {len(results)} файлів на диску {disk}. Результати збережено в {result_file}")

        else:
            print(f"❌ Помилка: диск {disk}:\\ не знайдено!")

    input("\n🔹 Натисніть Enter, для завершення роботи програми...")
