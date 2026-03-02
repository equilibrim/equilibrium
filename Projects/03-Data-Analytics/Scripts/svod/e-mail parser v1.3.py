import imaplib
import email
import os
import re
import argparse
import csv
import chardet
from email.header import decode_header
from datetime import datetime, timedelta
import logging

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class EmailOrganizationProcessor:
    def __init__(self, imap_server, email_address, password, organizations_file):
        """
        Инициализация обработчика писем с сортировкой по организациям
        """
        self.imap_server = imap_server
        self.email_address = email_address
        self.password = password
        self.mail = None
        
        # Основная папка для всех организаций
        self.base_folder = "Организации_и_письма"
        
        # Поддерживаемые форматы файлов
        self.supported_extensions = ['xlsx', 'pdf', 'docx', 'doc']
        
        # Словарь для отслеживания уже созданных папок организаций
        self.organizations_cache = {}
        
        # Загружаем список организаций
        self.organizations_mapping = self.load_organizations_mapping(organizations_file)
        
        # Создаем базовую папку
        os.makedirs(self.base_folder, exist_ok=True)
    
    def load_organizations_mapping(self, filepath):
        """Загрузка словаря соответствия 'ключ поиска' -> 'название папки' из файла"""
        mapping = {}
        if not os.path.exists(filepath):
            logging.error(f"❌ Файл списка организаций не найден: {filepath}")
            logging.info("Продолжаем без списка организаций, используя оригинальные имена.")
            return mapping

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip()]

            logging.info(f"📋 Загружено строк из {filepath}: {len(lines)}")
            for line in lines:
                # Формат: "ключ поиска | название папки"
                if '|' in line:
                    parts = line.split('|', 1) # Разбиваем на 2 части, на случай, если в названии есть '|'
                    if len(parts) == 2:
                        search_key = parts[0].strip()
                        folder_name = parts[1].strip()
                        if search_key and folder_name:
                            mapping[search_key] = folder_name
                        else:
                            logging.warning(f"Некорректная строка (пустой ключ или папка): {line}")
                    else:
                        logging.warning(f"Некорректная строка (нет '|'): {line}")
                else:
                    # Если строка не содержит '|', считаем её ключом (ключ = имя папки)
                    search_key = line.strip()
                    if search_key:
                         mapping[search_key] = search_key
            logging.info(f"✅ Загружено {len(mapping)} соответствий организаций.")
            
            # Сохраняем загруженные соответствия для отладки
            debug_file = os.path.join(self.base_folder, "отладка_организаций.txt")
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write("ЛОГИКА НАЗВАНИЯ ФАЙЛОВ:\n")
                f.write("Формат в файле: 'КЛЮЧ_ПОИСКА | НАЗВАНИЕ_ПАПКИ'\n")
                f.write("Или просто 'КЛЮЧ_ПОИСКА' (ключ = имя папки)\n")
                f.write("="*80 + "\n")
                f.write("СПИСОК КЛЮЧЕЙ ДЛЯ ПОИСКА ОРГАНИЗАЦИИ:\n")
                for search_key, folder_name in sorted(mapping.items()):
                     f.write(f"'{search_key}' -> '{folder_name}'\n")
            logging.info(f"Отладочная информация сохранена в: {debug_file}")
            
            return mapping
        except Exception as e:
            logging.error(f"❌ Ошибка загрузки файла организаций {filepath}: {e}")
            return {}

    def find_organization_name(self, sender_or_subject):
        """Поиск соответствия в заголовке (отправитель или тема) и возврат названия папки"""
        # Проверяем все ключи поиска в заголовке
        # Приводим к нижнему регистру для поиска
        search_text = sender_or_subject.lower()
        # Убираем разделители
        clean_search_text = re.sub(r'[_\-.]', ' ', search_text)

        for search_key, folder_name in self.organizations_mapping.items():
            # Ищем ключ в очищенном заголовке (без учета регистра и разделителей)
            if search_key.lower() in clean_search_text:
                return folder_name
        return None # Не найдено соответствие

    def clean_organization_name(self, name):
        """Очистка и нормализация названия организации (для случаев без списка)"""
        if not name:
            return ""
        
        # Преобразуем в строку если это не строка
        if not isinstance(name, str):
            name = str(name)
        
        # Декодируем если это bytes
        if isinstance(name, bytes):
            try:
                name = name.decode('utf-8')
            except:
                name = name.decode('cp1251', errors='ignore')
        
        # Удаляем лишние символы и слова
        patterns_to_remove = [
            r'["\']',  # Кавычки
            r'<[^>]+>',  # Email в угловых скобках
            r'\([^)]+\)',  # Текст в скобках
            r'\[[^\]]+\]',  # Текст в квадратных скобках
            r'\b(?:от|с|по|у|для|на)\b',  # Предлоги
            r'\s+',
        ]
        
        cleaned = name.strip()
        
        # Применяем паттерны очистки
        for pattern in patterns_to_remove:
            if pattern == r'\s+':
                cleaned = re.sub(pattern, ' ', cleaned)
            else:
                cleaned = re.sub(pattern, ' ', cleaned, flags=re.IGNORECASE)
        
        # Удаляем общие окончания компаний
        company_endings = [
            r'\s+(?:ооо|зао|ао|пао|ип|нко|кфх|llc|inc|ltd|gmbh)\b\.?$',
            r'\s+(?:company|corporation|limited|group|holding)\b\.?$',
        ]
        
        for ending in company_endings:
            cleaned = re.sub(ending, '', cleaned, flags=re.IGNORECASE)
        
        # Удаляем лишние пробелы
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        
        # Если название слишком длинное, обрезаем но сохраняем ключевые слова
        if len(cleaned) > 80:
            # Ищем сокращения и важные слова
            important_words = ['рб', 'рн', 'буз', 'црб', 'гбуз', 'мбуз', 'кдц', 'мсч']
            words = cleaned.split()
            
            # Оставляем первые 4 слова + любые важные слова
            important_found = []
            other_words = []
            
            for word in words:
                if word.lower() in important_words:
                    important_found.append(word)
                else:
                    other_words.append(word)
            
            # Берем первые 4 обычных слова + все важные
            selected_words = other_words[:4] + important_found
            if selected_words:
                cleaned = ' '.join(selected_words)
            else:
                cleaned = cleaned[:80]
        
        # Заменяем проблемные символы для Windows
        invalid_chars = r'<>:"/\|?*'
        for char in invalid_chars:
            cleaned = cleaned.replace(char, '_')
        
        # Удаляем точки в начале и конце
        cleaned = cleaned.strip('.')
        
        # Если после очистки пусто
        if not cleaned or len(cleaned) < 2:
            return "Неизвестная_организация"
        
        return cleaned
    
    def parse_email_date(self, date_str):
        """Парсинг даты из письма в datetime"""
        try:
            # Удаляем лишние пробелы
            date_str = date_str.strip()
            
            # Различные форматы дат в письмах
            date_formats = [
                '%a, %d %b %Y %H:%M:%S %z',
                '%d %b %Y %H:%M:%S %z',
                '%a, %d %b %Y %H:%M:%S',
                '%d %b %Y %H:%M:%S',
                '%Y-%m-%d %H:%M:%S',
                '%d.%m.%Y %H:%M:%S',
                '%d/%m/%Y %H:%M:%S',
            ]
            
            for fmt in date_formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except:
                    continue
            
            # Если не удалось распарсить, возвращаем текущую дату
            logging.warning(f"Не удалось распарсить дату: {date_str}")
            return datetime.now()
            
        except Exception as e:
            logging.error(f"Ошибка парсинга даты: {e}")
            return datetime.now()
    
    def format_date_for_folder(self, date_obj):
        """Форматирование даты для имени папки"""
        return date_obj.strftime("%Y-%m-%d_%H%M")
    
    def extract_organization_from_sender(self, sender):
        """Извлечение названия организации из поля отправителя"""
        try:
            # Декодируем отправителя
            decoded_sender = self.decode_header(sender)

            # Пытаемся найти организацию по списку
            matched_org_name = self.find_organization_name(decoded_sender)
            if matched_org_name:
                 # Используем имя из списка организаций для папки
                 safe_org_name_for_folder = self.clean_organization_name(matched_org_name)
                 # Возвращаем как (имя_для_папки, имя_для_файла)
                 # В данном случае они совпадают, но можно изменить логику
                 return safe_org_name_for_folder, matched_org_name

            # Если не нашли по списку, используем оригинальную логику
            # Разные форматы отправителя
            patterns = [
                r'"([^"]+)"\s*<[^>]+>',  # "Название" <email>
                r'([^<]+)\s*<[^>]+>',     # Название <email>
                r'(.+)\s+\([^)]*@[^)]*\)', # Название (email)
            ]
            
            for pattern in patterns:
                match = re.search(pattern, decoded_sender)
                if match:
                    candidate = match.group(1).strip()
                    cleaned = self.clean_organization_name(candidate)
                    if cleaned and cleaned != "Неизвестная_организация":
                        return cleaned, cleaned # имя_для_папки, имя_для_файла
            
            # Если ничего не нашли, пытаемся очистить всю строку
            fallback_name = self.clean_organization_name(decoded_sender)
            return fallback_name, fallback_name
            
        except Exception as e:
            logging.warning(f"Не удалось извлечь организацию из отправителя: {e}")
            return "Неизвестная_организация", "Неизвестная_организация"
    
    def get_organization_folder(self, organization_name, email_date):
        """Получение или создание папки организации и подпапки с датой"""
        # Очищаем имя организации для использования в пути
        safe_org_name = organization_name
        
        # Проверяем, есть ли уже папка для этой организации
        org_folder_path = None
        if organization_name in self.organizations_cache:
            org_folder_path = self.organizations_cache[organization_name]
        else:
            # Ищем существующую папку
            for item in os.listdir(self.base_folder):
                item_path = os.path.join(self.base_folder, item)
                if os.path.isdir(item_path):
                    # Проверяем схожесть имен (без учета регистра и окончаний)
                    item_clean = self.clean_organization_name(item)
                    if item_clean == organization_name:
                        org_folder_path = item_path
                        self.organizations_cache[organization_name] = org_folder_path
                        break
            
            # Если не нашли, создаем новую
            if not org_folder_path:
                # Создаем папку с оригинальным именем
                org_folder_path = os.path.join(self.base_folder, safe_org_name)
                
                # Если такая папка уже существует (но под другим названием)
                counter = 1
                original_path = org_folder_path
                while os.path.exists(org_folder_path):
                    org_folder_path = f"{original_path}_{counter}"
                    counter += 1
                
                os.makedirs(org_folder_path, exist_ok=True)
                self.organizations_cache[organization_name] = org_folder_path
                logging.info(f"Создана папка организации: {os.path.basename(org_folder_path)}")
        
        # Создаем подпапку с датой письма
        date_folder_name = self.format_date_for_folder(email_date)
        date_folder_path = os.path.join(org_folder_path, date_folder_name)
        
        # Если папка с такой датой уже существует, добавляем время с секундами
        counter = 1
        original_date_path = date_folder_path
        while os.path.exists(date_folder_path):
            date_folder_name = email_date.strftime("%Y-%m-%d_%H%M%S")
            if counter > 1:
                date_folder_name = f"{date_folder_name}_{counter}"
            date_folder_path = os.path.join(org_folder_path, date_folder_name)
            counter += 1
        
        os.makedirs(date_folder_path, exist_ok=True)
        
        return org_folder_path, date_folder_path, os.path.basename(org_folder_path), date_folder_name
    
    def decode_header(self, header):
        """Декодирование заголовка с исправлением кодировки"""
        if not header:
            return ""
        
        try:
            if isinstance(header, bytes):
                # Пытаемся определить кодировку
                result = chardet.detect(header)
                encoding = result['encoding'] if result['encoding'] else 'utf-8'
                try:
                    return header.decode(encoding)
                except:
                    return header.decode('utf-8', errors='ignore')
            
            decoded_parts = decode_header(header)
            result_parts = []
            
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        try:
                            result_parts.append(part.decode(encoding))
                        except:
                            try:
                                result_parts.append(part.decode('utf-8'))
                            except:
                                result_parts.append(part.decode('cp1251', errors='ignore'))
                    else:
                        try:
                            result_parts.append(part.decode('utf-8'))
                        except:
                            result_parts.append(part.decode('cp1251', errors='ignore'))
                else:
                    result_parts.append(str(part))
            
            return ''.join(result_parts)
            
        except Exception as e:
            logging.warning(f"Ошибка декодирования заголовка: {e}")
            return str(header)
    
    def get_email_body(self, msg):
        """Получение текста письма"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            result = chardet.detect(payload)
                            encoding = result['encoding'] if result['encoding'] else 'utf-8'
                            body = payload.decode(encoding, errors='ignore')
                            break
                    except Exception as e:
                        logging.warning(f"Ошибка декодирования тела письма: {e}")
                        continue
        else:
            try:
                payload = msg.get_payload(decode=True)
                if payload:
                    result = chardet.detect(payload)
                    encoding = result['encoding'] if result['encoding'] else 'utf-8'
                    body = payload.decode(encoding, errors='ignore')
            except Exception as e:
                logging.warning(f"Ошибка получения тела письма: {e}")
        
        return body
    
    def save_email_metadata(self, date_folder_path, email_data, organization):
        """Сохранение метаданных письма"""
        try:
            # Создаем текстовый файл с информацией о письме
            metadata_file = os.path.join(date_folder_path, "информация_о_письме.txt")
            
            with open(metadata_file, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("ИНФОРМАЦИЯ О ПИСЬМЕ\n")
                f.write("=" * 60 + "\n\n")
                
                f.write(f"Организация: {organization}\n")
                f.write(f"Дата письма: {email_data.get('date', 'Неизвестно')}\n")
                f.write(f"Тема: {email_data.get('subject', 'Без темы')}\n")
                f.write(f"Отправитель: {email_data.get('sender', 'Неизвестно')}\n")
                f.write(f"Дата обработки: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
                f.write(f"Количество вложений: {len(email_data.get('attachments', []))}\n\n")
                
                # Сохраняем первые 500 символов тела письма
                body = email_data.get('body', '')
                if body:
                    f.write("Текст письма (начало):\n")
                    f.write("-" * 40 + "\n")
                    f.write(body[:500])
                    if len(body) > 500:
                        f.write("\n... [текст обрезан]")
                    f.write("\n")
            
            # Также сохраняем в общий CSV файл организации
            org_folder_path = os.path.dirname(date_folder_path)
            csv_file = os.path.join(org_folder_path, "все_письма.csv")
            
            csv_data = {
                'Дата_письма': email_data.get('date', ''),
                'Дата_папки': os.path.basename(date_folder_path),
                'Тема': email_data.get('subject', ''),
                'Отправитель': email_data.get('sender', ''),
                'Вложений': len(email_data.get('attachments', [])),
                'Дата_обработки': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            file_exists = os.path.isfile(csv_file)
            with open(csv_file, 'a', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=csv_data.keys())
                if not file_exists:
                    writer.writeheader()
                writer.writerow(csv_data)
                
        except Exception as e:
            logging.error(f"Ошибка сохранения метаданных письма: {e}")
    
    def connect(self):
        """Подключение к почтовому серверу"""
        try:
            logging.info(f"Подключение к {self.imap_server}...")
            self.mail = imaplib.IMAP4_SSL(self.imap_server)
            self.mail.login(self.email_address, self.password)
            logging.info("✓ Успешное подключение!")
            return True
        except Exception as e:
            logging.error(f"✗ Ошибка подключения: {e}")
            return False
    
    def disconnect(self):
        """Отключение от сервера"""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
                logging.info("Отключение от сервера")
            except:
                pass
    
    def process_emails(self, days=7):
        """Основная обработка писем"""
        if not self.connect():
            return
        
        try:
            self.mail.select('INBOX')
            
            # Формируем дату для поиска
            since_date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
            
            # Ищем все письма за период
            result, data = self.mail.search(None, f'SINCE {since_date}')
            if result != 'OK':
                logging.error("Ошибка поиска писем")
                return
            
            email_ids = data[0].split()
            logging.info(f"Найдено писем за {days} дней: {len(email_ids)}")
            
            processed_count = 0
            files_saved = 0
            
            for i, email_id in enumerate(email_ids, 1):
                try:
                    logging.info(f"[{i}/{len(email_ids)}] Обработка письма...")
                    
                    # Получаем письмо
                    result, data = self.mail.fetch(email_id, '(RFC822)')
                    if result != 'OK':
                        continue
                    
                    # Парсим письмо
                    raw_email = data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # Извлекаем информацию
                    email_date_str = msg.get("Date", "")
                    email_date = self.parse_email_date(email_date_str)
                    
                    email_data = {
                        'date': email_date_str,
                        'date_obj': email_date,
                        'subject': self.decode_header(msg.get("Subject", "Без темы")),
                        'sender': self.decode_header(msg.get("From", "")),
                        'body': self.get_email_body(msg),
                        'attachments': []
                    }
                    
                    # Собираем вложения
                    for part in msg.walk():
                        if part.get_content_disposition() == 'attachment':
                            filename = part.get_filename()
                            if filename:
                                decoded_filename = self.decode_header(filename)
                                file_ext = os.path.splitext(decoded_filename)[1].lower().replace('.', '')
                                
                                if file_ext in self.supported_extensions:
                                    content = part.get_payload(decode=True)
                                    email_data['attachments'].append({
                                        'filename': decoded_filename,
                                        'content': content,
                                        'extension': file_ext
                                    })
                    
                    # Если есть нужные вложения, обрабатываем
                    if email_data['attachments']:
                        # Определяем организацию (имя для папки, имя для файла)
                        org_name_for_folder, org_name_for_file = self.extract_organization_from_sender(email_data['sender'])
                        
                        # Получаем пути для сохранения
                        org_folder_path, date_folder_path, org_name_actual, date_folder_name = self.get_organization_folder(
                            org_name_for_folder, email_date # Используем имя для папки
                        )
                        
                        # Сохраняем метаданные письма
                        self.save_email_metadata(date_folder_path, email_data, org_name_for_folder)
                        
                        # Сохраняем файлы в папку с датой
                        for attachment in email_data['attachments']:
                            # Создаем новое имя файла: [имя_организации_из_списка]_[оригинальное_имя_без_расширения].[расширение]
                            original_name_no_ext, original_ext = os.path.splitext(attachment['filename'])
                            
                            # Используем имя организации из списка (org_name_for_file) для начала имени файла
                            new_filename = f"{org_name_for_file}_{original_name_no_ext}{original_ext}"

                            # Создаем безопасное имя файла
                            safe_filename = re.sub(r'[^\w\-.]', '_', new_filename)
                            safe_filename = safe_filename[:150] # Ограничиваем общую длину имени файла
                        
                            # Добавляем индекс если нужно
                            filepath = os.path.join(date_folder_path, safe_filename)
                            
                            counter = 1
                            base_name, ext = os.path.splitext(filepath)
                            while os.path.exists(filepath):
                                filepath = f"{base_name}_{counter}{ext}"
                                counter += 1
                            
                            # Сохраняем файл
                            with open(filepath, 'wb') as f:
                                f.write(attachment['content'])
                            
                            files_saved += 1
                            logging.info(f"  ✓ Сохранен: {org_name_actual}/{date_folder_name}/{os.path.basename(filepath)}")
                        
                        processed_count += 1
                        logging.info(f"  Письмо сохранено в: {org_name_actual}/{date_folder_name}")
                    
                except Exception as e:
                    logging.error(f"Ошибка обработки письма: {e}")
                    continue
            
            # Генерируем отчет
            self.generate_report(processed_count, files_saved)
            
        except Exception as e:
            logging.error(f"Ошибка при обработке писем: {e}")
        finally:
            self.disconnect()
    
    def generate_report(self, processed_emails, saved_files):
        """Генерация отчета"""
        report_file = os.path.join(self.base_folder, "отчет_обработки.txt")
        
        # Собираем статистику
        org_stats = {}
        for item in os.listdir(self.base_folder):
            item_path = os.path.join(self.base_folder, item)
            if os.path.isdir(item_path) and item != "отчет_обработки.txt":
                # Подсчитываем папки с письмами
                email_folders = [d for d in os.listdir(item_path) 
                               if os.path.isdir(os.path.join(item_path, d))]
                
                # Подсчитываем файлы
                total_files = 0
                for email_folder in email_folders:
                    email_folder_path = os.path.join(item_path, email_folder)
                    files = [f for f in os.listdir(email_folder_path) 
                           if os.path.isfile(os.path.join(email_folder_path, f))]
                    # Исключаем файл информации о письме
                    files = [f for f in files if not f.endswith('информация_о_письме.txt')]
                    total_files += len(files)
                
                org_stats[item] = {
                    'письма': len(email_folders),
                    'файлы': total_files,
                    'папки_писем': email_folders
                }
        
        # Записываем отчет
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("ОТЧЕТ ОБ ОБРАБОТКЕ ЭЛЕКТРОННОЙ ПОЧТЫ\n")
            f.write("=" * 80 + "\n\n")
            
            f.write(f"Дата генерации отчета: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Файл списка организаций: {getattr(self, 'organizations_mapping_file', 'Не указан')}\n")
            f.write(f"Количество загруженных соответствий: {len(self.organizations_mapping)}\n")
            f.write(f"Обработано писем: {processed_emails}\n")
            f.write(f"Сохранено файлов: {saved_files}\n")
            f.write(f"Организаций: {len(org_stats)}\n\n")
            
            f.write("СТАТИСТИКА ПО ОРГАНИЗАЦИЯМ:\n")
            f.write("=" * 80 + "\n")
            
            # Сортируем организации по количеству писем
            sorted_orgs = sorted(org_stats.items(), 
                               key=lambda x: x[1]['письма'], 
                               reverse=True)
            
            for org_name, stats in sorted_orgs:
                f.write(f"\n🏢 {org_name}:\n")
                f.write(f"   📧 Писем: {stats['письма']}\n")
                f.write(f"   📁 Файлов: {stats['файлы']}\n")
                
                # Перечисляем даты писем
                if stats['папки_писем']:
                    f.write("   📅 Даты писем:\n")
                    for date_folder in sorted(stats['папки_писем'])[:10]:  # Показываем первые 10
                        f.write(f"      • {date_folder}\n")
                    if len(stats['папки_писем']) > 10:
                        f.write(f"      ... и еще {len(stats['папки_писем']) - 10} писем\n")
                
                f.write("-" * 40 + "\n")
            
            f.write("\n" + "=" * 80 + "\n")
            f.write("СТРУКТУРА ПАПОК:\n")
            f.write("=" * 80 + "\n\n")
            f.write("организации_и_письма/\n")
            f.write("├── [Название организации 1]/\n")
            f.write("│   ├── 2024-01-15_1430/        (папка с датой письма)\n")
            f.write("│   │   ├── 2_ГБУЗ_Акимовкая_ЦРБ_отчет1.xlsx\n") # Пример нового имени файла
            f.write("│   │   ├── 2_ГБУЗ_Акимовкая_ЦРБ_документ2.pdf\n")
            f.write("│   │   └── информация_о_письме.txt\n")
            f.write("│   ├── 2024-01-18_0920/\n")
            f.write("│   │   └── ...\n")
            f.write("│   └── все_письма.csv          (список всех писем)\n")
            f.write("├── [Название организации 2]/\n")
            f.write("│   └── ...\n")
            f.write("└── отчет_обработки.txt\n")
            
            f.write("\n" + "=" * 80 + "\n")
            f.write("ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО!\n")
            f.write("=" * 80 + "\n")
        
        logging.info(f"Отчет сохранен: {report_file}")


def main():
    parser = argparse.ArgumentParser(
        description='Обработка почты с сортировкой по организациям и датам писем'
    )
    parser.add_argument('--days', type=int, default=7, 
                       help='Количество дней для обработки (по умолчанию: 7)')
    parser.add_argument('--server', type=str, default='imap.mail.ru',
                       help='IMAP сервер (по умолчанию: imap.mail.ru)')
    parser.add_argument('--org-file', type=str, default='Список организаций.txt',
                       help='Файл со списком организаций (по умолчанию: Список организаций.txt)')
    
    args = parser.parse_args()
    
    print("=" * 70)
    print("📧 ОБРАБОТЧИК ПОЧТЫ - СОРТИРОВКА ПО ОРГАНИЗАЦИЯМ И ДАТАМ")
    print("=" * 70)
    print(f"Период: последние {args.days} дней")
    print(f"Сервер: {args.server}")
    print(f"Файл организаций: {args.org_file}")
    print("Форматы файлов: XLSX, PDF, DOCX, DOC")
    print("=" * 70)
    
    # Запрос учетных данных
    email_address = input("\nВведите email адрес: ").strip()
    password = input("Введите пароль: ").strip()
    
    # Создаем процессор
    processor = EmailOrganizationProcessor(
        imap_server=args.server,
        email_address=email_address,
        password=password,
        organizations_file=args.org_file # Передаем файл организаций
    )
    
    try:
        # Запускаем обработку
        processor.process_emails(days=args.days)
        
    except KeyboardInterrupt:
        print("\n\n⚠️ Программа прервана пользователем")
    except Exception as e:
        logging.error(f"Критическая ошибка: {e}")
    
    # Пауза перед закрытием
    input("\nНажмите Enter для выхода...")


if __name__ == "__main__":
    main()