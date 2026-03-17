<<<<<<< HEAD
# Интеграция API (Черновик)

## Прототип на FastAPI

```python
"""Апи"""
from fastapi import FastAPI
import json
import requests

app = FastAPI()

=======
"""Апи"""
from fastapi import FastAPI
import json

app = FastAPI()

# Дизайн системы
# Схема данных
# API

>>>>>>> 53265070ff5d362fb38377424e7448472008f56d
@app.get("/items/")
async def read_items():
    return [{"name": "Foo"}, {"name": "Bar"}]

@app.post("/items/")
async def create_item(item: dict):
    return item

<<<<<<< HEAD
def get_data_from_api(url, token):
    """
    Получает данные из API по переданному URL и ключу
=======

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
>>>>>>> 53265070ff5d362fb38377424e7448472008f56d
    """
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    return json.loads(response.text)
<<<<<<< HEAD
```
=======
>>>>>>> 53265070ff5d362fb38377424e7448472008f56d


# Варианты реализации интеграции
## Идеи интеграции

* Интеграция с системами управления версиями
* Интеграция с инструментами аналитики
* Интеграция с сервисами общения
* Интеграция с базами данных
