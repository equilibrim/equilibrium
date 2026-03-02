[[e-mail sorter v3]]
import os
import re
import shutil
from datetime import datetime
import openpyxl
import logging
from pathlib import Path
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class ReportSorter:
    def __init__(self, source_folder, output_folder, report_names_file, interactive=False):
        self.source_folder = source_folder
        self.output_folder = output_folder
        self.report_names_file = report_names_file
        self.interactive = interactive
        os.makedirs(output_folder, exist_ok=True)
        # Основные форматы
        self.supported_formats = ['.xlsx', '.xls', '.pdf', '.docx', '.doc']
        # Словари для хранения: {ключ_поиска: (название_папки, тип_поиска)}
        # тип_поиска: 'content' или 'filename'
        self.search_to_folder = {}
        self.found_folders = set()
        # Статистика
        self.stats = {
            'total_files': 0,
            'processed': 0,
            'sorted': 0,
            'not_found': 0,
            'errors': 0,
            'moved': 0,
            'interactive_choices': 0,
            'exact_matches': 0,
            'name_matches': 0,
            'new_keys_added': 0
        }
        # Для хранения неотсортированных файлов
        self.unsorted_files = []
        self.all_files_original = []
        # Лог файл
        self.log_file = os.path.join(output_folder, "детальный_лог.txt")
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write(f"Лог сортировки - {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write("="*60 + "\n")

    def log_detail(self, message):
        """Запись детального лога"""
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")

    def extract_organization_from_path(self, file_path, rel_path):
        """Извлечение названия организации из пути к файлу"""
        try:
            # Путь может быть: Организации_и_письма/Название_организации/2024-01-15_1430/файл.xlsx
            parts = rel_path.split(os.sep)
            if len(parts) >= 2:
                # Берем название организации (первая папка после базовой)
                org_name = parts[0]
                # Очищаем название
                org_name = re.sub(r'[<>:"/\\|?*]', '_', org_name)
                org_name = org_name.strip('_')
                # Если название слишком длинное, сокращаем
                if len(org_name) > 30:
                    # Берем первые 20 символов + последние 5
                    org_name = org_name[:20] + "..." + org_name[-5:]
                return org_name if org_name else "Неизвестно"
            # Если путь простой, пытаемся извлечь из имени файла
            filename = os.path.basename(file_path)
            # Ищем паттерны в имени файла
            patterns = [
                r'от\s+([^_\-.]+)',  # "от Название"
                r'([А-Я][а-я]+)\s+отчет',  # "Название отчет"
                r'([А-Я]+[\w\s]+)_отчет',  # "НАЗВАНИЕ_отчет"
            ]
            for pattern in patterns:
                match = re.search(pattern, filename, re.IGNORECASE)
                if match:
                    org_name = match.group(1).strip()
                    org_name = re.sub(r'[<>:"/\\|?*]', '_', org_name)
                    return org_name[:30]
            return "Неизвестно"
        except Exception:
            return "Неизвестно"

    def load_report_names(self):
        """Загрузка ключей поиска и названий папок из файла"""
        print(f"\n📋 Загрузка настроек из: {self.report_names_file}")
        if not os.path.exists(self.report_names_file):
            print(f"❌ Файл не найден: {self.report_names_file}")
            return False

        try:
            with open(self.report_names_file, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip()]

            print(f"📄 Загружено строк: {len(lines)}")
            for line in lines:
                # Формат: "ключ | папка | тип" или "ключ | папка" (по умолчанию content) или просто "ключ"
                parts = line.split('|', 2) # Разбиваем максимум на 3 части
                if len(parts) == 3:
                    search_key = parts[0].strip()
                    folder_name = parts[1].strip()
                    search_type = parts[2].strip().lower()
                    if search_key and folder_name and search_type in ['content', 'filename']:
                        self.search_to_folder[search_key] = (folder_name, search_type)
                elif len(parts) == 2:
                    search_key = parts[0].strip()
                    folder_name = parts[1].strip()
                    if search_key and folder_name:
                        self.search_to_folder[search_key] = (folder_name, 'content') # По умолчанию content
                else: # len(parts) == 1
                    # Просто ключ (ключ = имя папки, поиск в содержимом)
                    search_key = line.strip()
                    if search_key:
                        self.search_to_folder[search_key] = (search_key, 'content')

            print(f"✅ Загружено ключей поиска: {len(self.search_to_folder)}")
            print(f"✅ Будут созданы папки: {len(set([v[0] for v in self.search_to_folder.values()]))}")
            # Сохраняем настройки для отладки
            debug_file = os.path.join(self.output_folder, "настройки_поиска.txt")
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write("НАСТРОЙКИ ПОИСКА И СОРТИРОВКИ:\n")
                f.write("="*80 + "\n")
                f.write("Формат: 'КЛЮЧ_ПОИСКА | НАЗВАНИЕ_ПАПКИ | ТИП_ПОИСКА'\n")
                f.write("ТИП_ПОИСКА: 'content' или 'filename'\n")
                f.write("ИЛИ просто 'КЛЮЧ_ПОИСКА' (ключ = имя папки, поиск в содержимом)\n")
                f.write("="*80 + "\n")
                f.write("📋 СПИСОК КЛЮЧЕЙ ДЛЯ ПОИСКА:\n")
                for search_key, (folder_name, search_type) in sorted(self.search_to_folder.items()):
                    f.write(f"\n🔍 Ищем: '{search_key}' (тип: {search_type})")
                    if search_key != folder_name:
                        f.write(f" → 📁 Папка: '{folder_name}'")
                    f.write("\n")
            return True
        except Exception as e:
            print(f"❌ Ошибка загрузки: {e}")
            return False

    def save_report_names(self):
        """Сохранение ключей поиска в файл"""
        try:
            with open(self.report_names_file, 'w', encoding='utf-8') as f:
                for search_key, (folder_name, search_type) in self.search_to_folder.items():
                    if search_key == folder_name and search_type == 'content':
                        f.write(f"{search_key}\n")
                    else:
                        f.write(f"{search_key} | {folder_name} | {search_type}\n")
            print(f"✅ Настройки сохранены в файл: {self.report_names_file}")
            return True
        except Exception as e:
            print(f"❌ Ошибка сохранения настроек: {e}")
            return False

    def search_exact_in_excel(self, file_path, filename):
        """ТОЧНЫЙ поиск ключей в содержимом Excel файла"""
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            # Собираем ВЕСЬ текст из ВСЕХ листов
            all_text_lines = []
            for sheet_name in wb.sheetnames:  # Все листы
                ws = wb[sheet_name]
                # Читаем все строки до 500 и колонки до 20
                for row in ws.iter_rows(min_row=1, max_row=500, min_col=1, max_col=20, values_only=True):
                    row_texts = []
                    for cell in row:
                        if cell:
                            cell_text = str(cell).strip()
                            if cell_text:
                                row_texts.append(cell_text)
                    if row_texts:
                        # Сохраняем строки как есть (для поиска целых строк)
                        row_line = ' '.join(row_texts)
                        all_text_lines.append(row_line)
            wb.close()

            if not all_text_lines:
                return None

            # Ищем ТОЧНЫЕ совпадения с ключами, учитывая тип поиска
            for search_key, (folder_name, search_type) in self.search_to_folder.items():
                if search_type == 'content':
                    # Ищем точное вхождение ключа в любой строке
                    for line in all_text_lines:
                        if search_key in line:
                            return folder_name
            return None
        except Exception as e:
            self.log_detail(f"Ошибка чтения Excel {filename}: {e}")
            return None

    def search_exact_in_pdf(self, file_path, filename):
        """ТОЧНЫЙ поиск ключей в содержимом PDF"""
        try:
            import PyPDF2
            with open(file_path, 'rb') as f:
                try:
                    pdf_reader = PyPDF2.PdfReader(f)
                    # Читаем ВСЕ страницы
                    pdf_lines = []
                    for page_num in range(len(pdf_reader.pages)):
                        text = pdf_reader.pages[page_num].extract_text()
                        if text:
                            # Разбиваем на строки
                            lines = text.split('\n')
                            for line in lines:
                                line_clean = line.strip()
                                if line_clean:
                                    pdf_lines.append(line_clean)

                    if pdf_lines:
                        # Ищем ТОЧНЫЕ совпадения, учитывая тип поиска
                        for search_key, (folder_name, search_type) in self.search_to_folder.items():
                            if search_type == 'content':
                                for line in pdf_lines:
                                    if search_key in line:
                                        return folder_name
                except Exception as pdf_error:
                    self.log_detail(f"Ошибка PDF {filename}: {pdf_error}")
                    return None
        except ImportError:
            return None
        except Exception:
            return None

    def search_in_filename(self, filename):
        """Поиск ключей в имени файла, учитывая тип поиска"""
        # Убираем расширение и очищаем имя
        name_without_ext = os.path.splitext(filename)[0]
        # Приводим к нижнему регистру для поиска
        clean_name = re.sub(r'[_\-.]', ' ', name_without_ext.lower())

        for search_key, (folder_name, search_type) in self.search_to_folder.items():
            # Ищем ключ в имени файла (без учета регистра и разделителей) ТОЛЬКО если тип 'filename'
            if search_type == 'filename':
                 if search_key.lower() in clean_name:
                    return folder_name
        return None

    def identify_report_type(self, file_path):
        """Поиск ТОЛЬКО в содержимом файлов (оригинальная логика)"""
        filename = os.path.basename(file_path)
        file_ext = os.path.splitext(filename)[1].lower()

        # Excel файлы
        if file_ext in ['.xlsx', '.xls']:
            return self.search_exact_in_excel(file_path, filename)
        # PDF файлы
        elif file_ext == '.pdf':
            return self.search_exact_in_pdf(file_path, filename)
        # Другие форматы - только в содержимом
        elif file_ext in ['.docx', '.doc']:
            return None
        return None

    def identify_report_type_with_filename(self, file_path):
        """Поиск в содержимом файлов И в именах файлов (для ресортировки)"""
        filename = os.path.basename(file_path)
        file_ext = os.path.splitext(filename)[1].lower()

        # Сначала проверяем имя файла
        folder_name = self.search_in_filename(filename)
        if folder_name:
            self.stats['name_matches'] += 1  # Увеличиваем счётчик при совпадении по имени
            return folder_name

        # Если в имени файла ничего не нашли, ищем в содержимом
        # Excel файлы
        if file_ext in ['.xlsx', '.xls']:
            return self.search_exact_in_excel(file_path, filename)
        # PDF файлы
        elif file_ext == '.pdf':
            return self.search_exact_in_pdf(file_path, filename)
        # Другие форматы - только в содержимом
        elif file_ext in ['.docx', '.doc']:
            return None
        return None

    def get_interactive_choice(self, filename, file_ext, file_path, organization):
        """Интерактивный выбор для нераспознанных файлов"""
        print(f"\n{'='*60}")
        print(f"❓ ФАЙЛ НЕ РАСПОЗНАН: {filename}")
        print(f"   Полный путь: {file_path}")
        print(f"   Организация: {organization}")
        print(f"   Формат: {file_ext}")
        print(f"{'-'*60}")
        # Показываем существующие папки
        existing_folders = sorted(list(self.found_folders))
        if existing_folders:
            print("Существующие папки:")
            for i, folder in enumerate(existing_folders[:20], 1):
                print(f"  {i:2}. {folder}")
            if len(existing_folders) > 20:
                print(f"  ... и еще {len(existing_folders) - 20} папок")

        print("\nВыберите действие:")
        print("  1. Создать новую папку")
        if existing_folders:
            print("  2. Выбрать из существующих папок")
        print("  3. Поместить в 'НЕ_СОРТИРОВАННЫЕ'")
        print("  4. Пропустить файл (оставить на месте)")
        print("  5. Добавить новый ключ поиска в содержимое")
        print("  6. Добавить новый ключ поиска по имени файла")
        print("  7. Просмотреть содержимое файла")

        while True:
            choice = input("\nВаш выбор: ").strip()
            if choice == '1':
                folder_name = input("Введите название новой папки: ").strip()
                if folder_name:
                    return folder_name
                else:
                    print("Название папки не может быть пустым!")
            elif choice == '2' and existing_folders:
                try:
                    folder_num = int(input(f"Введите номер папки (1-{len(existing_folders)}): "))
                    if 1 <= folder_num <= len(existing_folders):
                        return existing_folders[folder_num - 1]
                    else:
                        print(f"Неверный номер! Введите от 1 до {len(existing_folders)}")
                except ValueError:
                    print("Введите число!")
            elif choice == '3':
                return "НЕ_СОРТИРОВАННЫЕ"
            elif choice == '4':
                return None
            elif choice == '5':
                # Добавляем новый ключ для поиска в содержимом
                result = self.add_new_search_key(file_path, filename, file_ext, search_type='content')
                if result:
                    return result
                else:
                    print("Продолжаем выбор папки для текущего файла...")
                    continue
            elif choice == '6':
                # Добавляем новый ключ для поиска по имени файла
                result = self.add_new_search_key(file_path, filename, file_ext, search_type='filename')
                if result:
                    return result
                else:
                    print("Продолжаем выбор папки для текущего файла...")
                    continue
            elif choice == '7':
                self.preview_file_content(file_path, file_ext)
                continue
            else:
                print("Неверный выбор! Введите 1, 2, 3, 4, 5, 6 или 7")

    def add_new_search_key(self, file_path, filename, file_ext, search_type='content'):
        """Добавление нового ключа поиска"""
        print(f"\n➕ ДОБАВЛЕНИЕ НОВОГО КЛЮЧА ПОИСКА ({'в содержимом' if search_type == 'content' else 'в имени файла'})")
        print(f"Файл: {filename}")

        # Сначала показываем содержимое файла для помощи
        print("\nСодержимое файла (первые 200 символов):")
        content_preview = self.get_file_preview(file_path, file_ext, max_chars=200)
        print(f"  {content_preview}")

        if search_type == 'content':
            # Логика для поиска в содержимом
            print("\nВы можете:")
            print("  1. Ввести текст вручную")
            print("  2. Выбрать текст из содержимого файла")
            choice = input("Ваш выбор (1 или 2): ").strip()
            search_key = ""
            if choice == '2':
                # Показываем больше содержимого для выбора
                print("\nСодержимое файла для выбора текста:")
                full_preview = self.get_file_preview(file_path, file_ext, max_chars=1000)
                lines = full_preview.split('\n')
                print("="*60)
                for i, line in enumerate(lines[:20], 1):
                    print(f"{i:2}. {line}")
                print("="*60)
                try:
                    line_num = int(input(f"\nВыберите номер строки (1-{min(20, len(lines))}): "))
                    if 1 <= line_num <= len(lines):
                        selected_line = lines[line_num-1]
                        print(f"\nВыбранная строка: '{selected_line}'")
                        # Предлагаем выбрать часть строки
                        print("\nВведите часть строки для использования как ключ поиска:")
                        print(f"  Строка: {selected_line}")
                        search_key = input("  Ключ поиска: ").strip()
                except (ValueError, IndexError):
                    print("Неверный выбор, вводите текст вручную.")
                    search_key = input("\nВведите текст для поиска в файлах: ").strip()
            else:
                search_key = input("\nВведите текст для поиска в файлах: ").strip()
        else:  # search_type == 'filename'
            # Логика для поиска в имени файла
            print("\nВведите текст, который должен содержаться в имени файла:")
            print(f"  Текущее имя: {filename}")
            search_key = input("  Ключ поиска в имени файла: ").strip()

        if not search_key:
            print("Ключ поиска не может быть пустым!")
            return None

        print("\nВведите название папки для этого ключа:")
        print("  1. Создать новую папку")
        print("  2. Выбрать из существующих")
        folder_choice_input = input("  Ваш выбор (1 или 2): ").strip()
        folder_name = ""

        existing_folders = sorted(list(self.found_folders))
        if folder_choice_input == '2' and existing_folders:
            print("\nСуществующие папки:")
            for i, folder in enumerate(existing_folders, 1):
                print(f"  {i:2}. {folder}")
            try:
                folder_num = int(input(f"\nВведите номер папки (1-{len(existing_folders)}): "))
                if 1 <= folder_num <= len(existing_folders):
                    folder_name = existing_folders[folder_num - 1]
                else:
                    print(f"Неверный номер! Введите от 1 до {len(existing_folders)}")
                    return None
            except ValueError:
                print("Введите число!")
                return None
        elif folder_choice_input == '1':
            folder_name = input("\nВведите название новой папки: ").strip()
        else:
            print("Неверный выбор!")
            return None

        if not folder_name:
            print("Название папки не может быть пустым!")
            return None

        # Добавляем новый ключ
        self.search_to_folder[search_key] = (folder_name, search_type)
        self.stats['new_keys_added'] += 1
        print(f"\n✅ Добавлен ключ поиска: '{search_key}' → папка '{folder_name}' (тип поиска: {search_type})")

        # Сохраняем настройки в файл
        self.save_report_names()

        # Автоматически ищем новый ключ в текущем файле (учитываем тип поиска)
        print(f"\n🔍 Поиск нового ключа в текущем файле...")
        found_folder = self.find_folder_by_newest_key(file_path, search_type)
        if found_folder:
            print(f"✅ Найден ключ! Файл будет перемещен в папку: '{found_folder}'")
            return found_folder
        else:
            print("⚠️  Ключ не найден в текущем файле.")

        # Предлагаем выполнить ресортировку неотсортированных файлов
        if self.unsorted_files:
            print(f"\n🔄 Обнаружено {len(self.unsorted_files)} неотсортированных файлов")
            rescan = input("Выполнить ресортировку неотсортированных файлов с новым ключом? (да/нет): ").strip().lower()
            if rescan == 'да':
                # Выполняем ресортировку
                sorted_count = self.rescan_unsorted_files()
                print(f"✅ Ресортировано {sorted_count} файлов с новым ключом")
                # Показываем статистику
                print(f"\n📊 После ресортировки:")
                print(f"   Всего отсортировано: {self.stats['sorted']}")
                print(f"   Осталось неотсортированных: {len(self.unsorted_files)}")

                # Проверяем, сортируется ли текущий файл после ресортировки
                found_folder_after_rescan = self.find_folder_by_newest_key(file_path, search_type)
                if found_folder_after_rescan:
                    print(f"✅ Текущий файл теперь сортируется в папку: '{found_folder_after_rescan}'")
                    return found_folder_after_rescan

        # Спрашиваем, как поступить с текущим файлом
        print(f"\nКак поступить с текущим файлом '{filename}'?")
        action = input("  Создать папку для ключа (1) или выбрать другую папку (2)? ")
        if action == '1':
            return folder_name
        else:
            return None

    def find_folder_by_newest_key(self, file_path, added_search_type):
        """Вспомогательная функция для поиска по типу только что добавленного ключа"""
        filename = os.path.basename(file_path)
        file_ext = os.path.splitext(filename)[1].lower()

        if added_search_type == 'filename':
             # Сначала проверяем имя файла
            folder_name = self.search_in_filename(filename)
            if folder_name:
                return folder_name
        elif added_search_type == 'content':
            # Проверяем содержимое
            if file_ext in ['.xlsx', '.xls']:
                return self.search_exact_in_excel(file_path, filename)
            elif file_ext == '.pdf':
                return self.search_exact_in_pdf(file_path, filename)
            elif file_ext in ['.docx', '.doc']:
                return None # или реализовать для docx/doc
        return None


    def get_file_preview(self, file_path, file_ext, max_chars=200):
        """Получение предпросмотра содержимого файла"""
        try:
            if file_ext in ['.xlsx', '.xls']:
                # Для Excel файлов
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                sheet = wb.active
                preview_lines = []
                for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                    row_data = [str(cell) for cell in row if cell]
                    if row_data:
                        preview_lines.append(f"Строка {i}: {' | '.join(row_data[:5])}")
                wb.close()
                return '\n'.join(preview_lines)
            elif file_ext == '.pdf':
                # Для PDF файлов
                try:
                    import PyPDF2
                    with open(file_path, 'rb') as f:
                        pdf_reader = PyPDF2.PdfReader(f)
                        text = pdf_reader.pages[0].extract_text()
                        return text[:max_chars] + ('...' if len(text) > max_chars else '')
                except Exception:
                    return "[Не удалось прочитать содержимое PDF]"
            else:
                return "[Просмотр содержимого недоступен для этого формата]"
        except Exception as e:
            return f"[Ошибка при чтении файла: {e}]"

    def preview_file_content(self, file_path, file_ext):
        """Предварительный просмотр содержимого файла"""
        try:
            print(f"\n📄 Просмотр содержимого файла:")
            print(f"   Путь: {file_path}")
            if file_ext in ['.xlsx', '.xls']:
                # Для Excel файлов показываем первые строки
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                sheet = wb.active
                print(f"   Лист: {sheet.title}")
                print(f"   Размер: {sheet.max_row} строк, {sheet.max_column} колонок")
                print("\nПервые 10 строк:")
                for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                    row_data = [str(cell)[:50] for cell in row if cell]
                    if row_data:
                        print(f"   {i:2}. {' | '.join(row_data)}")
                wb.close()
            elif file_ext == '.pdf':
                # Для PDF файлов пытаемся извлечь текст
                try:
                    import PyPDF2
                    with open(file_path, 'rb') as f:
                        pdf_reader = PyPDF2.PdfReader(f)
                        text = pdf_reader.pages[0].extract_text()
                        print(f"   Страниц: {len(pdf_reader.pages)}")
                        print("\nТекст первой страницы:")
                        lines = text.split('\n')
                        for i, line in enumerate(lines[:15], 1):
                            print(f"   {i:2}. {line[:80]}")
                except Exception:
                    print("   Не удалось прочитать содержимое PDF")
            else:
                print("   Просмотр содержимого недоступен для этого формата")
        except Exception as e:
            print(f"   Ошибка при просмотре файла: {e}")

    def rescan_unsorted_files(self):
        """Ресортировка неотсортированных файлов с новыми ключами, учитывая их тип."""
        print(f"\n🔄 РЕСОРТИРОВКА НЕОТСОРТИРОВАННЫХ ФАЙЛОВ")
        print(f"Неотсортированных файлов: {len(self.unsorted_files)}")
        print(f"Новых ключей поиска: {self.stats['new_keys_added']}")
        
        sorted_count = 0
        # Создаем копию списка для безопасной итерации
        unsorted_copy = self.unsorted_files.copy()
        
        for i, (file_path, rel_path, organization) in enumerate(unsorted_copy, 1):
            print(f"\n📋 Проверка файла {i}/{len(unsorted_copy)}")
            print(f"   Файл: {os.path.basename(file_path)}")
            print(f"   Организация: {organization}")

            # Проверяем, не был ли файл уже перемещен
            if not os.path.exists(file_path):
                print(f"   ⚠️  Файл уже перемещен, удаляем из списка")
                self.unsorted_files.remove((file_path, rel_path, organization))
                continue

            # --- ОСНОВНОЕ ИЗМЕНЕНИЕ ---
            # Используем универсальный метод, который проверяет как имя, так и содержимое,
            # но с учетом типа поиска, указанного в словаре.
            folder_name = self.identify_report_type_with_filename(file_path)

            if folder_name:
                print(f"   ✅ Найдено совпадение! Папка: '{folder_name}'")
                # Перемещаем файл
                if self.move_file_to_folder(file_path, folder_name, organization):
                    self.stats['sorted'] += 1
                    self.stats['not_found'] -= 1
                    self.unsorted_files.remove((file_path, rel_path, organization))
                    sorted_count += 1
                else:
                    print(f"   ❌ Ошибка перемещения файла")
            else:
                print(f"   ❌ Совпадений не найдено")

        return sorted_count

    def create_final_filename(self, original_filename, organization):
        """Создание окончательного имени файла с отправителем"""
        # Очищаем название организации
        safe_org = re.sub(r'[<>:"/\\|?*]', '_', organization)
        safe_org = safe_org.strip('_')
        # Если организация неизвестна, не добавляем префикс
        if safe_org == "Неизвестно" or not safe_org:
            return original_filename
        # Получаем расширение файла
        name_without_ext, ext = os.path.splitext(original_filename)
        # Формируем новое имя: [Организация]_оригинальное_имя.расширение
        # Но если файл уже начинается с этой организации, не дублируем
        if original_filename.lower().startswith(safe_org.lower() + '_'):
            return original_filename
        new_filename = f"{safe_org}_{original_filename}"
        # Ограничиваем длину (Windows ограничение - 260 символов)
        if len(new_filename) > 200:
            # Сокращаем имя файла, но оставляем организацию и расширение
            max_name_len = 200 - len(ext) - len(safe_org) - 2  # -2 для подчеркиваний
            if max_name_len > 10:
                name_part = name_without_ext[:max_name_len]
                new_filename = f"{safe_org}_{name_part}{ext}"
            else:
                # Если слишком длинно, оставляем только организацию и расширение
                new_filename = f"{safe_org}{ext}"
        return new_filename

    def move_file_to_folder(self, source_path, target_folder_name, organization):
        """ПЕРЕМЕЩЕНИЕ файла в целевую папку с добавлением отправителя в имя"""
        # Создаем безопасное имя папки
        safe_folder_name = re.sub(r'[<>:"/\\|?*]', '_', target_folder_name)
        safe_folder_name = safe_folder_name[:100].strip()
        # Создаем целевую папку
        target_dir = os.path.join(self.output_folder, safe_folder_name)
        os.makedirs(target_dir, exist_ok=True)
        # Добавляем в список папок
        self.found_folders.add(safe_folder_name)
        # Создаем новое имя файла с отправителем
        original_filename = os.path.basename(source_path)
        final_filename = self.create_final_filename(original_filename, organization)
        target_path = os.path.join(target_dir, final_filename)
        # Если файл уже существует, добавляем номер
        counter = 1
        base_name, ext = os.path.splitext(target_path)
        while os.path.exists(target_path):
            target_path = f"{base_name}_{counter}{ext}"
            counter += 1
        try:
            # ВАЖНО: ПЕРЕМЕЩАЕМ файл (не копируем!)
            shutil.move(source_path, target_path)
            self.stats['moved'] += 1
            # Логируем
            log_msg = f"  ПЕРЕМЕЩЕН в: {safe_folder_name}/{os.path.basename(target_path)}"
            if counter > 1:
                log_msg += f" (переименован с {original_filename})"
            # Если имя изменилось, показываем старое и новое
            if final_filename != original_filename:
                log_msg += f" [было: {original_filename}]"
            print(log_msg)
            self.log_detail(log_msg)
            return True
        except Exception as e:
            error_msg = f"  ❌ Ошибка перемещения {original_filename}: {e}"
            print(error_msg)
            self.log_detail(f"  Ошибка перемещения {original_filename}: {e}")
            self.stats['errors'] += 1
            return False

    def scan_all_files(self):
        """Сканирование всех файлов"""
        print(f"\n🔍 Сканирование папки: {self.source_folder}")
        all_files = []
        for root, dirs, files in os.walk(self.source_folder):
            for file in files:
                file_ext = os.path.splitext(file)[1].lower()
                if file_ext in self.supported_formats:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(root, self.source_folder)
                    all_files.append((file_path, rel_path))

        self.stats['total_files'] = len(all_files)
        print(f"✅ Найдено файлов: {self.stats['total_files']}")
        # Сохраняем оригинальный список файлов
        self.all_files_original = all_files.copy()
        return all_files

    def process_file(self, file_info):
        """Обработка одного файла"""
        file_path, rel_path = file_info
        try:
            self.stats['processed'] += 1
            current_num = self.stats['processed']
            total_files = self.stats['total_files']

            # Вывод прогресса
            if current_num % 50 == 0:
                print(f"📊 [{current_num:4}/{total_files:4}] "
                      f"Отсортировано: {self.stats['sorted']:4} | "
                      f"Точных совпадений: {self.stats['exact_matches']:4} | "
                      f"По имени: {self.stats['name_matches']:4} | "
                      f"Не найдено: {self.stats['not_found']:4}")

            filename = os.path.basename(file_path)

            # Извлекаем организацию из пути
            organization = self.extract_organization_from_path(file_path, rel_path)

            # ТОЛЬКО поиск в содержимом файла (оригинальная логика)
            folder_name = self.identify_report_type(file_path)

            if folder_name:
                self.stats['exact_matches'] += 1
                # Перемещаем файл с добавлением отправителя в имя
                if self.move_file_to_folder(file_path, folder_name, organization):
                    self.stats['sorted'] += 1
                    return (file_path, folder_name, True, "Успешно перемещен", organization)
                else:
                    return (file_path, None, False, "Ошибка перемещения", organization)
            else:
                # Файл не распознан
                if self.interactive:
                    # Добавляем в список неотсортированных для интерактивной обработки
                    self.unsorted_files.append((file_path, rel_path, organization))
                    self.stats['not_found'] += 1
                    return (file_path, None, False, "Ожидает интерактивной обработки", organization)
                else:
                    self.stats['not_found'] += 1
                    # Автоматически перемещаем в НЕ_СОРТИРОВАННЫЕ
                    if self.move_file_to_folder(file_path, "НЕ_СОРТИРОВАННЫЕ", organization):
                        return (file_path, "НЕ_СОРТИРОВАННЫЕ", True, "Перемещен в НЕ_СОРТИРОВАННЫЕ", organization)
                    else:
                        return (file_path, None, False, "Ошибка перемещения в НЕ_СОРТИРОВАННЫЕ", organization)
        except Exception as e:
            self.stats['errors'] += 1
            error_msg = f"Критическая ошибка обработки {file_path}: {e}"
            print(f"❌ {error_msg}")
            self.log_detail(error_msg)
            return (file_path, None, False, str(e), "Неизвестно")

    def process_interactive_files(self):
        """Обработка файлов в интерактивном режиме"""
        print(f"\n🔧 ИНТЕРАКТИВНЫЙ РЕЖИМ")
        print(f"Файлов для ручной обработки: {len(self.unsorted_files)}")
        print("="*60)
        # Создаем копию списка для безопасной итерации
        unsorted_copy = self.unsorted_files.copy()

        for i, (file_path, rel_path, organization) in enumerate(unsorted_copy, 1):
            # Проверяем, не был ли файл уже перемещен
            if not os.path.exists(file_path):
                print(f"\n⚠️  Файл уже перемещен, пропускаем")
                continue

            filename = os.path.basename(file_path)
            file_ext = os.path.splitext(filename)[1].lower()

            print(f"\n📋 Файл {i}/{len(unsorted_copy)}: {filename}")
            print(f"   Организация: {organization}")

            # Интерактивный выбор
            folder_choice = self.get_interactive_choice(filename, file_ext, file_path, organization)

            if folder_choice:
                self.stats['interactive_choices'] += 1
                if self.move_file_to_folder(file_path, folder_choice, organization):
                    self.stats['sorted'] += 1
                    self.stats['not_found'] -= 1  # Уменьшаем счетчик нераспознанных
                    # Удаляем из списка неотсортированных
                    if (file_path, rel_path, organization) in self.unsorted_files:
                        self.unsorted_files.remove((file_path, rel_path, organization))
                else:
                    # Если ошибка перемещения, оставляем в исходной папке
                    print(f"  ⚠️  Файл оставлен в исходной папке: {file_path}")
            else:
                print(f"  ⚠️  Файл пропущен: {filename}")
                # Файл остается в исходной папке и в списке неотсортированных

    def process_all_files(self, max_workers=4):
        """Обработка всех файлов"""
        if not self.load_report_names():
            return False

        all_files = self.scan_all_files()
        if not all_files:
            print("⚠️ Файлы не найдены!")
            return False

        print(f"\n🚀 Начинаем обработку {len(all_files)} файлов...")
        print("="*60)
        print("⚠️  ВНИМАНИЕ: Ищем ТОЛЬКО в содержимом файлов (при первичной обработке)")
        print("⚠️  Имена файлов игнорируются на первом этапе!")
        print("⚠️  К именам файлов будет добавлен отправитель")
        print("⚠️  Файлы ПЕРЕМЕЩАЮТСЯ (не копируются)!")
        print("="*60)

        # Создаем папку для неотсортированных
        unsorted_folder = os.path.join(self.output_folder, "НЕ_СОРТИРОВАННЫЕ")
        os.makedirs(unsorted_folder, exist_ok=True)
        self.found_folders.add("НЕ_СОРТИРОВАННЫЕ")

        # Обработка файлов
        results = []

        if self.interactive:
            # В интерактивном режиме используем только один поток
            print("\n🔄 Обработка файлов в однопоточном режиме (интерактивный режим)...")
            for file_info in all_files:
                file_path, rel_path = file_info
                filename = os.path.basename(file_path)
                # Обновляем прогресс
                self.stats['processed'] += 1
                if self.stats['processed'] % 10 == 0:
                    print(f"📊 [{self.stats['processed']:4}/{len(all_files):4}] "
                          f"Отсортировано: {self.stats['sorted']:4} "
                          f"Неотсортировано: {len(self.unsorted_files):4}")

                organization = self.extract_organization_from_path(file_path, rel_path)

                # Сначала пытаемся автоматически определить ТОЛЬКО по содержимому
                folder_name = self.identify_report_type(file_path)

                if folder_name:
                    # Автоматическое перемещение
                    self.stats['exact_matches'] += 1
                    if self.move_file_to_folder(file_path, folder_name, organization):
                        self.stats['sorted'] += 1
                        results.append((file_path, folder_name, True, "Успешно перемещен", organization))
                    else:
                        results.append((file_path, None, False, "Ошибка перемещения", organization))
                else:
                    # Добавляем в список для интерактивной обработки
                    self.unsorted_files.append((file_path, rel_path, organization))
                    results.append((file_path, None, False, "Ожидает интерактивной обработки", organization))
                    self.stats['not_found'] += 1

            # Обработка файлов в интерактивном режиме
            if self.unsorted_files:
                self.process_interactive_files()
        else:
            # Неинтерактивный режим - используем многопоточность
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_file = {executor.submit(self.process_file, file_info): file_info
                                  for file_info in all_files}
                for future in as_completed(future_to_file):
                    try:
                        result = future.result()
                        results.append(result)
                    except Exception as e:
                        error_msg = f"Ошибка в потоке: {e}"
                        print(f"❌ {error_msg}")
                        self.log_detail(error_msg)

        # Генерация отчета
        self.generate_report(results)
        return True

    def generate_report(self, results):
        """Генерация итогового отчета"""
        report_file = os.path.join(self.output_folder, "ИТОГОВЫЙ_ОТЧЕТ.txt")

        # Группируем результаты
        report_stats = {}
        organizations_used = set()
        for file_path, folder_name, success, message, organization in results:
            if success and folder_name:
                report_stats[folder_name] = report_stats.get(folder_name, 0) + 1
            if organization and organization != "Неизвестно":
                organizations_used.add(organization)

        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("ИТОГОВЫЙ ОТЧЕТ СОРТИРОВКИ\n")
            f.write("="*80 + "\n")
            f.write(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Исходная папка: {self.source_folder}\n")
            f.write(f"Выходная папка: {self.output_folder}\n")
            f.write(f"Файл с настройками: {self.report_names_file}\n")
            f.write(f"Интерактивный режим: {'Да' if self.interactive else 'Нет'}\n")
            if self.interactive:
                f.write(f"Новых ключей добавлено: {self.stats['new_keys_added']}\n")
            f.write("⚠️  РЕЖИМ ПОИСКА (при первичной обработке): ТОЛЬКО В СОДЕРЖИМОМ ФАЙЛОВ\n")
            f.write("⚠️  ИМЕНА ФАЙЛОВ ИГНОРИРУЮТСЯ!\n")
            f.write("⚠️  К именам файлов добавлен отправитель (если известен)\n")
            f.write("⚠️  ФАЙЛЫ ПЕРЕМЕЩАЮТСЯ, А НЕ КОПИРУЮТСЯ!\n")
            f.write("="*80 + "\n")
            f.write("СТАТИСТИКА\n")
            f.write("="*80 + "\n")
            f.write(f"Всего файлов: {self.stats['total_files']}\n")
            f.write(f"Обработано: {self.stats['processed']}\n")
            f.write(f"Успешно перемещено: {self.stats['moved']}\n")
            f.write(f"Точных совпадений в содержимом: {self.stats['exact_matches']}\n")
            f.write(f"Совпадений по имени файла (после добавления ключей): {self.stats['name_matches']}\n")  # Добавлено
            if self.interactive:
                f.write(f"Интерактивных выборов: {self.stats['interactive_choices']}\n")
                f.write(f"Добавлено новых ключей: {self.stats['new_keys_added']}\n")
            f.write(f"Не распознано: {self.stats['not_found']}\n")
            f.write(f"Ошибок: {self.stats['errors']}\n")
            if self.interactive and self.unsorted_files:
                f.write(f"⚠️  Осталось неотсортированных файлов: {len(self.unsorted_files)}\n")

            # Детализация по папкам
            if report_stats:
                f.write("="*80 + "\n")
                f.write("РАСПРЕДЕЛЕНИЕ ФАЙЛОВ ПО ПАПКАМ\n")
                f.write("="*80 + "\n")
                # Сортируем по количеству файлов
                sorted_stats = sorted(report_stats.items(), key=lambda x: x[1], reverse=True)
                for folder_name, count in sorted_stats:
                    f.write(f"📁 {folder_name}: {count} файл(ов)\n")

            # Информация об организациях
            if organizations_used:
                f.write("\n" + "="*80 + "\n")
                f.write("ИСПОЛЬЗОВАННЫЕ ОРГАНИЗАЦИИ-ОТПРАВИТЕЛИ\n")
                f.write("="*80 + "\n")
                for org in sorted(organizations_used):
                    f.write(f"🏢 {org}\n")

            # Файлы, оставшиеся в исходной папке
            remaining_files = [(file_path, message) for file_path, folder_name, success, message, organization
                               in results if not success or not folder_name or message == "Ожидает интерактивной обработки"]
            if remaining_files:
                f.write("\n" + "="*80 + "\n")
                f.write("ФАЙЛЫ, ОСТАВШИЕСЯ В ИСХОДНОЙ ПАПКЕ\n")
                f.write("="*80 + "\n")
                for file_path, message in remaining_files[:50]:  # Ограничиваем вывод
                    filename = os.path.basename(file_path)
                    f.write(f"❌ {filename}: {message}\n")
                if len(remaining_files) > 50:
                    f.write(f"\n... и еще {len(remaining_files) - 50} файлов\n")

            f.write("\n" + "="*80 + "\n")
            f.write("ВНИМАНИЕ\n")
            f.write("="*80 + "\n")
            f.write("1. Файлы были ПЕРЕМЕЩЕНЫ из исходной папки\n")
            f.write("2. Исходные файлы больше не существуют в исходном расположении\n")
            f.write("3. Для отмены операции потребуется восстановление из бэкапа\n")
            f.write("4. Всегда делайте бэкап перед запуском сортировки!\n")

            if self.interactive and self.stats['new_keys_added'] > 0:
                f.write(f"✅ Добавлено {self.stats['new_keys_added']} новых ключей поиска\n")
                f.write("   Ключи сохранены в файле настроек\n")
                f.write("   Их можно использовать при следующем запуске\n")

            f.write("✅ Сортировка завершена!\n")

        print(f"\n📊 Отчет сохранен: {report_file}")
        print(f"📝 Подробный лог: {self.log_file}")

        # Вывод краткой статистики в консоль
        print("\n" + "="*60)
        print("ИТОГИ:")
        print(f"📁 Всего файлов: {self.stats['total_files']}")
        print(f"✅ Перемещено: {self.stats['moved']}")
        print(f"🎯 Точных совпадений в содержимом: {self.stats['exact_matches']}")
        print(f"🎯 Совпадений по имени файла (после добавления ключей): {self.stats['name_matches']}")  # Добавлено
        if self.interactive:
            print(f"👤 Интерактивных выборов: {self.stats['interactive_choices']}")
            print(f"➕ Новых ключей добавлено: {self.stats['new_keys_added']}")
            print(f"❓ Осталось неотсортированных: {len(self.unsorted_files)}")
        else:
            print(f"❓ Не распознано/оставлено: {self.stats['not_found']}")
        print(f"⚠️  Ошибок: {self.stats['errors']}")
        print("="*60)

def main():
    """Главная функция"""
    parser = argparse.ArgumentParser(description='Сортировка отчетов по содержимому файлов')
    parser.add_argument('--source', required=True, help='Исходная папка с файлами')
    parser.add_argument('--output', required=True, help='Выходная папка для сортировки')
    parser.add_argument('--config', required=True, help='Файл с названиями отчетов и ключами поиска')
    parser.add_argument('--interactive', action='store_true', help='Интерактивный режим')
    parser.add_argument('--workers', type=int, default=4, help='Количество потоков (по умолчанию: 4)')

    args = parser.parse_args()

    print("="*80)
    print("📁 СОРТИРОВЩИК ОТЧЕТОВ ПО СОДЕРЖИМОМУ ФАЙЛОВ")
    print("="*80)
    print(f"Исходная папка: {args.source}")
    print(f"Выходная папка: {args.output}")
    print(f"Файл настроек: {args.config}")
    print(f"Интерактивный режим: {'Да' if args.interactive else 'Нет'}")
    print(f"Потоков обработки: {args.workers}")
    print("="*80)
    print("⚠️  ВНИМАНИЕ: Файлы будут ПЕРЕМЕЩЕНЫ, а не скопированы!")
    print("⚠️  Рекомендуется сделать резервную копию перед запуском!")
    print("="*80)

    # Предупреждение
    if args.interactive:
        print("\n🔄 ИНТЕРАКТИВНЫЙ РЕЖИМ ВКЛЮЧЕН")
        print("Для каждого нераспознанного файла будет запрошено действие.")
        print("Можно добавлять новые ключи поиска (в содержимом или в имени файла) и выполнять ресортировку.")
        confirm = input("\nПродолжить? (да/НЕТ): ").strip().lower()
        if confirm != 'да':
            print("Отменено пользователем.")
            return

    # Проверка исходной папки
    if not os.path.exists(args.source):
        print(f"❌ Исходная папка не существует: {args.source}")
        return

    if not os.path.exists(args.config):
        print(f"❌ Файл настроек не существует: {args.config}")
        return

    # Создание экземпляра сортировщика
    sorter = ReportSorter(
        source_folder=args.source,
        output_folder=args.output,
        report_names_file=args.config,
        interactive=args.interactive
    )

    # Запуск обработки
    try:
        success = sorter.process_all_files(max_workers=args.workers if not args.interactive else 1)
        if success:
            print("\n✅ Сортировка завершена успешно!")
            print(f"\n📁 Результаты в папке: {args.output}")
            # Показываем созданные папки
            if os.path.exists(args.output):
                folders = [d for d in os.listdir(args.output)
                           if os.path.isdir(os.path.join(args.output, d))]
                if folders:
                    print(f"\n📂 Создано папок: {len(folders)}")
                    print("Основные папки:")
                    for folder in sorted(folders)[:10]:  # Показываем первые 10
                        print(f"  📁 {folder}")
                    if len(folders) > 10:
                        print(f"  ... и еще {len(folders) - 10} папок")
        else:
            print("\n❌ Сортировка завершена с ошибками!")

    except KeyboardInterrupt:
        print("\n⚠️  Процесс прерван пользователем!")
        print("⚠️  Частичные результаты сохранены.")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()