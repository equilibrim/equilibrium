import os

# Настройка путей (скрипт лежит в Projects/01-CNC-Laser/Scripts/)
current_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(current_dir)
photos_dir = os.path.join(project_dir, "Photos")

# Словарь соответствия: "Старое имя" -> "Новое имя"
renaming_rules = {
    "photo_2026-02-28_11-55-49.jpg": "Arduino_Mega_2560.jpg",
    # Сюда буду добавлять следующие фото по мере их появления
}

print(f"📂 Сканирую папку: {photos_dir}")

for old_name, new_name in renaming_rules.items():
    old_path = os.path.join(photos_dir, old_name)
    new_path = os.path.join(photos_dir, new_name)
    
    if os.path.exists(old_path):
        try:
            os.rename(old_path, new_path)
            print(f"✅ Переименовано: {old_name} -> {new_name}")
        except Exception as e:
            print(f"❌ Ошибка при переименовании {old_name}: {e}")
    elif os.path.exists(new_path):
        print(f"ℹ️ Файл уже переименован: {new_name}")
    else:
        print(f"⚠️ Исходный файл не найден: {old_name}")
import os

# Настройка путей (скрипт лежит в Projects/01-CNC-Laser/Scripts/)
current_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(current_dir)
photos_dir = os.path.join(project_dir, "Photos")

# Словарь соответствия: "Старое имя" -> "Новое имя"
renaming_rules = {
    "photo_2026-02-28_11-55-49.jpg": "Arduino_Mega_2560.jpg",
    # Сюда буду добавлять следующие фото по мере их появления
}

print(f"📂 Сканирую папку: {photos_dir}")

for old_name, new_name in renaming_rules.items():
    old_path = os.path.join(photos_dir, old_name)
    new_path = os.path.join(photos_dir, new_name)
    
    if os.path.exists(old_path):
        try:
            os.rename(old_path, new_path)
            print(f"✅ Переименовано: {old_name} -> {new_name}")
        except Exception as e:
            print(f"❌ Ошибка при переименовании {old_name}: {e}")
    elif os.path.exists(new_path):
        print(f"ℹ️ Файл уже переименован: {new_name}")
    else:
        print(f"⚠️ Исходный файл не найден: {old_name}")