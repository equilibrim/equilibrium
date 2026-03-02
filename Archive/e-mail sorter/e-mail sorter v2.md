
import os
import re
import shutil
from datetime import datetime
import openpyxl
import logging
from pathlib import Path
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed

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
        
        # Словари для хранения: {ключ_поиска: название_папки}
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
            'exact_matches': 0
        }
        
        # Лог файл
        self.log_file = os.path.join(output_folder, "детальный_лог.txt")
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write(f"Лог сортировки - {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write("="*60 + "\n")
    
    def log_detail(self, message):
        """Запись детального лога"""
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
    
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
                # Формат: "ключ | папка" или просто "ключ"
                if '|' in line:
                    parts = line.split('|', 1)
                    if len(parts) == 2:
                        search_key = parts[0].strip()
                        folder_name = parts[1].strip()
                        
                        if search_key and folder_name:
                            self.search_to_folder[search_key] = folder_name
                    
                else:
                    # Просто ключ (ключ = имя папки)
                    search_key = line.strip()
                    if search_key:
                        self.search_to_folder[search_key] = search_key
            
            print(f"✅ Загружено ключей поиска: {len(self.search_to_folder)}")
            print(f"✅ Будут созданы папки: {len(set(self.search_to_folder.values()))}")
            
            # Сохраняем настройки для отладки
            debug_file = os.path.join(self.output_folder, "настройки_поиска.txt")
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write("НАСТРОЙКИ ПОИСКА И СОРТИРОВКИ:\n")
                f.write("="*80 + "\n")
                f.write("Формат: 'КЛЮЧ_ПОИСКА | НАЗВАНИЕ_ПАПКИ'\n")
                f.write("ИЛИ просто 'КЛЮЧ_ПОИСКА' (ключ = имя папки)\n")
                f.write("="*80 + "\n\n")
                
                f.write("📋 СПИСОК КЛЮЧЕЙ ДЛЯ ПОИСКА В СОДЕРЖИМОМ ФАЙЛОВ:\n")
                for search_key, folder_name in sorted(self.search_to_folder.items()):
                    f.write(f"\n🔍 Ищем: '{search_key}'")
                    if search_key != folder_name:
                        f.write(f" → 📁 Папка: '{folder_name}'")
                    f.write("\n")
            
            return True
            
        except Exception as e:
            print(f"❌ Ошибка загрузки: {e}")
            return False
    
    def search_exact_in_excel(self, file_path, filename):
        """ТОЧНЫЙ поиск ключей в содержимом Excel файла"""
        try:
            print(f"  🔍 Точный поиск в Excel: {filename}")
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
            
            # Ищем ТОЧНЫЕ совпадения с ключами
            for search_key, folder_name in self.search_to_folder.items():
                # Ищем точное вхождение ключа в любой строке
                for line in all_text_lines:
                    # ТОЧНОЕ совпадение (игнорируем только пробелы в начале/конце)
                    if search_key in line:
                        print(f"  ✅ ТОЧНОЕ СОВПАДЕНИЕ: '{search_key}' → папка '{folder_name}'")
                        print(f"     Найдено в строке: '{line[:100]}...'")
                        self.stats['exact_matches'] += 1
                        return folder_name
            
            return None
            
        except Exception as e:
            self.log_detail(f"Ошибка чтения Excel {filename}: {e}")
            return None
    
    def search_exact_in_pdf(self, file_path, filename):
        """ТОЧНЫЙ поиск ключей в содержимом PDF"""
        try:
            print(f"  📄 Точный поиск в PDF: {filename}")
            
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
                            # Ищем ТОЧНЫЕ совпадения
                            for search_key, folder_name in self.search_to_folder.items():
                                for line in pdf_lines:
                                    if search_key in line:
                                        print(f"  ✅ ТОЧНОЕ СОВПАДЕНИЕ в PDF: '{search_key}' → '{folder_name}'")
                                        print(f"     Найдено в строке: '{line[:100]}...'")
                                        self.stats['exact_matches'] += 1
                                        return folder_name
                    
                    except Exception as pdf_error:
                        self.log_detail(f"Ошибка PDF {filename}: {pdf_error}")
                        
            except ImportError:
                print(f"  ⚠️  PyPDF2 не установлен, пропускаем PDF: {filename}")
            
            return None
            
        except Exception:
            return None
    
    def identify_report_type(self, file_path):
        """ТОЛЬКО поиск в содержимом файлов (игнорируем имена файлов)"""
        filename = os.path.basename(file_path)
        file_ext = os.path.splitext(filename)[1].lower()
        
        self.log_detail(f"ТОЧНЫЙ поиск в файле: {filename}")
        
        # Excel файлы
        if file_ext in ['.xlsx', '.xls']:
            return self.search_exact_in_excel(file_path, filename)
        
        # PDF файлы
        elif file_ext == '.pdf':
            if self.interactive:
                print(f"\n📄 PDF файл: {filename}")
                return None
            else:
                return self.search_exact_in_pdf(file_path, filename)
        
        # Другие форматы - только в содержимом
        elif file_ext in ['.docx', '.doc']:
            print(f"  ⚠️  Файлы Word (.docx/.doc) пока не поддерживаются для поиска")
            return None
        
        return None
    
    def get_interactive_choice(self, filename, file_ext):
        """Интерактивный выбор для нераспознанных файлов"""
        print(f"\n{'='*60}")
        print(f"❓ ФАЙЛ НЕ РАСПОЗНАН В СОДЕРЖИМОМ: {filename}")
        print(f"   Формат: {file_ext}")
        print(f"{'-'*60}")
        
        # Показываем существующие папки
        existing_folders = sorted(self.found_folders)
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
        print("  5. Добавить новый ключ поиска")
        
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
                search_key = input("Введите точный текст для поиска в файлах: ").strip()
                folder_name = input("Введите название папки для этого ключа: ").strip()
                if search_key and folder_name:
                    # Добавляем новый ключ
                    self.search_to_folder[search_key] = folder_name
                    print(f"✅ Добавлен ключ поиска: '{search_key}' → папка '{folder_name}'")
                    
                    # Пытаемся найти этот ключ в текущем файле
                    if file_ext in ['.xlsx', '.xls']:
                        result = self.search_exact_in_excel(file_path, filename)
                        if result:
                            return result
                    elif file_ext == '.pdf':
                        result = self.search_exact_in_pdf(file_path, filename)
                        if result:
                            return result
                    
                    print("⚠️  Ключ не найден в текущем файле, создаем папку")
                    return folder_name
                else:
                    print("Ключ и название папки не могут быть пустыми!")
            
            else:
                print("Неверный выбор! Введите 1, 2, 3, 4 или 5")
    
    def move_file_to_folder(self, source_path, target_folder_name):
        """Перемещение файла в целевую папку"""
        # Создаем безопасное имя папки
        safe_folder_name = re.sub(r'[<>:"/\\|?*]', '_', target_folder_name)
        safe_folder_name = safe_folder_name[:100].strip()
        
        # Создаем целевую папку
        target_dir = os.path.join(self.output_folder, safe_folder_name)
        os.makedirs(target_dir, exist_ok=True)
        
        # Добавляем в список папок
        self.found_folders.add(safe_folder_name)
        
        # Имя файла остается исходным
        filename = os.path.basename(source_path)
        target_path = os.path.join(target_dir, filename)
        
        # Если файл уже существует, добавляем номер
        counter = 1
        base_name, ext = os.path.splitext(target_path)
        while os.path.exists(target_path):
            target_path = f"{base_name}_{counter}{ext}"
            counter += 1
        
        try:
            # ПЕРЕМЕЩАЕМ файл (не копируем!)
            shutil.move(source_path, target_path)
            self.stats['moved'] += 1
            
            # Логируем
            log_msg = f"  Перемещен в: {safe_folder_name}/{filename}"
            if counter > 1:
                log_msg += f" (переименован в {os.path.basename(target_path)})"
            self.log_detail(log_msg)
            
            return True
            
        except Exception as e:
            self.log_detail(f"  Ошибка перемещения {filename}: {e}")
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
                      f"Не найдено: {self.stats['not_found']:4}")
            
            filename = os.path.basename(file_path)
            
            # ТОЛЬКО поиск в содержимом файла
            folder_name = self.identify_report_type(file_path)
            
            if folder_name:
                # Перемещаем файл
                if self.move_file_to_folder(file_path, folder_name):
                    self.stats['sorted'] += 1
                    return (file_path, folder_name, True, "Успешно перемещен")
                else:
                    return (file_path, None, False, "Ошибка перемещения")
            
            else:
                # Файл не распознан в содержимом
                if self.interactive:
                    file_ext = os.path.splitext(filename)[1].lower()
                    folder_choice = self.get_interactive_choice(filename, file_ext)
                    
                    if folder_choice:
                        self.stats['interactive_choices'] += 1
                        if self.move_file_to_folder(file_path, folder_choice):
                            self.stats['sorted'] += 1
                            return (file_path, folder_choice, True, "Перемещен по выбору пользователя")
                        else:
                            self.stats['not_found'] += 1
                            return (file_path, None, False, "Ошибка перемещения по выбору")
                    else:
                        self.stats['not_found'] += 1
                        return (file_path, None, False, "Пропущен пользователем")
                
                else:
                    self.stats['not_found'] += 1
                    # Автоматически перемещаем в НЕ_СОРТИРОВАННЫЕ
                    if self.move_file_to_folder(file_path, "НЕ_СОРТИРОВАННЫЕ"):
                        return (file_path, "НЕ_СОРТИРОВАННЫЕ", True, "Перемещен в НЕ_СОРТИРОВАННЫЕ")
                    else:
                        return (file_path, None, False, "Ошибка перемещения в НЕ_СОРТИРОВАННЫЕ")
                
        except Exception as e:
            self.stats['errors'] += 1
            self.log_detail(f"Критическая ошибка обработки {file_path}: {e}")
            return (file_path, None, False, str(e))
    
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
        print("⚠️  ВНИМАНИЕ: Ищем ТОЛЬКО в содержимом файлов")
        print("⚠️  Имена файлов игнорируются!")
        print("="*60)
        
        # Создаем папку для неотсортированных
        unsorted_folder = os.path.join(self.output_folder, "НЕ_СОРТИРОВАННЫЕ")
        os.makedirs(unsorted_folder, exist_ok=True)
        self.found_folders.add("НЕ_СОРТИРОВАННЫЕ")
        
        # Обработка файлов
        results = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {executor.submit(self.process_file, file_info): file_info 
                            for file_info in all_files}
            
            for future in as_completed(future_to_file):
                try:
                    result = future.result()
                    results.append(result)
                except Exception as e:
                    self.log_detail(f"Ошибка в потоке: {e}")
        
        # Генерация отчета
        self.generate_report(results)
        
        return True
    
    def generate_report(self, results):
        """Генерация итогового отчета"""
        report_file = os.path.join(self.output_folder, "ИТОГОВЫЙ_ОТЧЕТ.txt")
        
        # Группируем результаты
        report_stats = {}
        for file_path, folder_name, success, message in results:
            if success and folder_name:
                report_stats[folder_name] = report_stats.get(folder_name, 0) + 1
        
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("ИТОГОВЫЙ ОТЧЕТ СОРТИРОВКИ\n")
            f.write("="*80 + "\n\n")
            
            f.write(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Исходная папка: {self.source_folder}\n")
            f.write(f"Выходная папка: {self.output_folder}\n")
            f.write(f"Файл с настройками: {self.report_names_file}\n")
            f.write(f"Интерактивный режим: {'Да' if self.interactive else 'Нет'}\n\n")
            
            f.write("⚠️  РЕЖИМ ПОИСКА: ТОЛЬКО В СОДЕРЖИМОМ ФАЙЛОВ ⚠️\n")
            f.write("   Имена файлов игнорируются!\n\n")
            
            f.write("СТАТИСТИКА ОБРАБОТКИ:\n")
            f.write("-"*80 + "\n")
            f.write(f"Всего файлов: {self.stats['total_files']}\n")
            f.write(f"Обработано: {self.stats['processed']}\n")
            f.write(f"Отсортировано: {self.stats['sorted']}\n")
            f.write(f"Не распознано: {self.stats['not_found']}\n")
            f.write(f"Ошибок: {self.stats['errors']}\n")
            f.write(f"Перемещено: {self.stats['moved']}\n")
            f.write(f"Точных совпадений в содержимом: {self.stats['exact_matches']}\n")
            
            if self.interactive:
                f.write(f"Интерактивных выборов: {self.stats['interactive_choices']}\n")
            
            # Проценты
            if self.stats['processed'] > 0:
                sorted_pct = (self.stats['sorted'] / self.stats['processed']) * 100
                unknown_pct = (self.stats['not_found'] / self.stats['processed']) * 100
                exact_pct = (self.stats['exact_matches'] / self.stats['sorted'] * 100) if self.stats['sorted'] > 0 else 0
                f.write(f"\nУспешно отсортировано: {sorted_pct:.1f}%\n")
                f.write(f"Не распознано: {unknown_pct:.1f}%\n")
                f.write(f"Найдено точных совпадений: {exact_pct:.1f}% от отсортированных\n")
            
            # Распределение по папкам
            if report_stats:
                f.write("\nРАСПРЕДЕЛЕНИЕ ПО ПАПКАМ:\n")
                f.write("-"*80 + "\n")
                
                sorted_stats = sorted(report_stats.items(), key=lambda x: x[1], reverse=True)
                total_sorted = sum(report_stats.values())
                
                for folder, count in sorted_stats:
                    pct = (count / total_sorted * 100) if total_sorted > 0 else 0
                    f.write(f"{folder[:70]:70} : {count:4} файлов ({pct:.1f}%)\n")
            
            # Связи ключ → папка
            f.write("\n" + "="*80 + "\n")
            f.write("КЛЮЧИ ДЛЯ ПОИСКА В СОДЕРЖИМОМ:\n")
            f.write("-"*80 + "\n")
            
            for search_key, folder_name in sorted(self.search_to_folder.items()):
                f.write(f"🔍 Ищем: '{search_key}' → 📁 Папка: '{folder_name}'\n")
            
            # Примеры строк где найдены ключи
            f.write("\n" + "="*80 + "\n")
            f.write("ПРИМЕР НАЙДЕННЫХ СОВПАДЕНИЙ (из лога):\n")
            f.write("-"*80 + "\n")
            
            # Читаем примеры из лога
            try:
                with open(self.log_file, 'r', encoding='utf-8') as log_f:
                    log_lines = log_f.readlines()
                    match_lines = [line for line in log_lines if 'ТОЧНОЕ СОВПАДЕНИЕ' in line]
                    
                    for line in match_lines[:10]:  # Первые 10 примеров
                        f.write(line)
                    
                    if len(match_lines) > 10:
                        f.write(f"\n... и еще {len(match_lines) - 10} совпадений\n")
            except:
                f.write("Не удалось прочитать примеры из лога\n")
        
        # Вывод в консоль
        print("\n" + "="*80)
        print("✅ СОРТИРОВКА ЗАВЕРШЕНА!")
        print("="*80)
        print(f"Всего файлов: {self.stats['total_files']}")
        print(f"Отсортировано: {self.stats['sorted']} ({sorted_pct:.1f}%)")
        print(f"Точных совпадений в содержимом: {self.stats['exact_matches']} ({exact_pct:.1f}%)")
        print(f"Не распознано: {self.stats['not_found']} ({unknown_pct:.1f}%)")
        print(f"Перемещено: {self.stats['moved']}")
        
        if self.stats['not_found'] > 0:
            print(f"\n⚠️  {self.stats['not_found']} файлов в папке 'НЕ_СОРТИРОВАННЫЕ'")
            print("   Возможные причины:")
            print("   1. Ключ не найден в содержимом файла")
            print("   2. Файл поврежден или имеет другой формат")
            print("   3. Ключ написан с ошибками или отличается от текста в файле")
        
        print(f"\n📁 Результаты в папке: {self.output_folder}")
        print(f"📄 Подробный отчет: {report_file}")
        print(f"📋 Настройки поиска: {self.output_folder}/настройки_поиска.txt")
        print(f"📋 Детальный лог: {self.log_file}")


def main():
    parser = argparse.ArgumentParser(description='Сортировка отчетов по точному совпадению в содержимом')
    parser.add_argument('source_folder', help='Папка с исходными файлами')
    parser.add_argument('output_folder', help='Папка для отсортированных файлов')
    parser.add_argument('report_names_file', help='Файл с настройками поиска (txt)')
    parser.add_argument('--workers', type=int, default=4, help='Количество потоков (по умолчанию: 4)')
    parser.add_argument('--interactive', action='store_true', help='Интерактивный режим для нераспознанных файлов')
    
    args = parser.parse_args()
    
    print("="*80)
    print("🎯 ТОЧНЫЙ СОРТИРОВЩИК ОТЧЕТОВ")
    print("="*80)
    print("⚠️  РЕЖИМ: ТОЛЬКО ПОИСК В СОДЕРЖИМОМ ФАЙЛОВ")
    print("⚠️  Имена файлов игнорируются!")
    print("="*80)
    print("ФОРМАТ НАСТРОЕК: 'ТОЧНЫЙ_ТЕКСТ_ДЛЯ_ПОИСКА | НАЗВАНИЕ_ПАПКИ'")
    print("   или просто 'ТОЧНЫЙ_ТЕКСТ_ДЛЯ_ПОИСКА' (ключ = имя папки)")
    print(f"Режим: {'ИНТЕРАКТИВНЫЙ' if args.interactive else 'АВТОМАТИЧЕСКИЙ'}")
    print(f"Исходная папка: {args.source_folder}")
    print(f"Выходная папка: {args.output_folder}")
    print(f"Файл с настройками: {args.report_names_file}")
    print(f"Потоков: {args.workers}")
    print("="*80)
    
    # Проверка
    if not os.path.exists(args.source_folder):
        print(f"❌ Папка не существует: {args.source_folder}")
        return
    
    if not os.path.exists(args.report_names_file):
        print(f"❌ Файл не существует: {args.report_names_file}")
        return
    
    # Запуск
    sorter = ReportSorter(
        args.source_folder, 
        args.output_folder, 
        args.report_names_file,
        interactive=args.interactive
    )
    
    try:
        success = sorter.process_all_files(max_workers=args.workers)
        
        if success:
            print("\n🎉 Сортировка успешно завершена!")
            print("\n📝 КЛЮЧЕВЫЕ МОМЕНТЫ:")
            print("   • Ищет ТОЧНЫЙ текст в содержимом файлов")
            print("   • Имена файлов игнорируются")
            print("   • 'Приложение 1' и 'Приложение 10' - разные ключи")
            print("   • Пробелы в начале/конце ключа игнорируются")
        else:
            print("\n⚠️  Сортировка завершена с ошибками")
        
    except KeyboardInterrupt:
        print("\n\n⏹️  Прервано пользователем")
        print("Файлы, которые уже были перемещены, остаются в целевых папках")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
    
    input("\nНажмите Enter для выхода...")


if __name__ == "__main__":
    # Установите библиотеки если нужно:
    # pip install openpyxl PyPDF2
    main()