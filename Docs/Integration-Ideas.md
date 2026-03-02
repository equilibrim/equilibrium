"""Апи"""
from fastapi import FastAPI
import json

app = FastAPI()

# Дизайн системы
# Схема данных
# API

@app.get("/items/")
async def read_items():
    return [{"name": "Foo"}, {"name": "Bar"}]

@app.post("/items/")
async def create_item(item: dict):
    return item


# Основные функции
def get_data_from_api(url, token):
    """
        Получает данные из API по переданному URL и ключу
        :param url: Адрес для запроса
        :type url: str
        :param token: Ключ для авторизации
        :type token: str
        :return: Результат выполнения API-requests
        :rtype: dict
    """
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    return json.loads(response.text)


# Варианты реализации интеграции
## Идеи интеграции

* Интеграция с системами управления версиями
* Интеграция с инструментами аналитики
* Интеграция с сервисами общения
* Интеграция с базами данных
