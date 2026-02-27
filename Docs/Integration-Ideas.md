# Интеграция API (Черновик)

## Прототип на FastAPI

```python
"""Апи"""
from fastapi import FastAPI
import json
import requests

app = FastAPI()

@app.get("/items/")
async def read_items():
    return [{"name": "Foo"}, {"name": "Bar"}]

@app.post("/items/")
async def create_item(item: dict):
    return item

def get_data_from_api(url, token):
    """
    Получает данные из API по переданному URL и ключу
    """
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    return json.loads(response.text)
```


# Варианты реализации интеграции
## Идеи интеграции

* Интеграция с системами управления версиями
* Интеграция с инструментами аналитики
* Интеграция с сервисами общения
* Интеграция с базами данных
