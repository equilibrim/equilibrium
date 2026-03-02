# 🧠 Equilibrium

> [!quote] Vision
> «Единая система знаний — основа инженерного мастерства»
> — *В. & А.*

---

## 🕹️ Control Center

> [!tip] **Active Context**
> Перед началом работы и перед закрытием — сверься с контекстом:
> 👉 **[[Docs/SESSION_CONTEXT|🟢 SESSION CONTEXT]]**

### Быстрые действия

```button
name 💡 Новая идея
type command
action Templater: Create new note from template
class primary
```

```button
name 📝 Новый проект
type command
action Templater: Create new note from template
class primary
```

```button
name ✅ Задачи
type link
action obsidian://open?file=Tasks
class secondary
```

```button
name ♻️ Синхронизировать
type command
action Git: Commit-and-sync
class secondary
```

---

## 🔭 Обзор системы

### Активные проекты

> [!example]- Список проектов (развернуть)
> ```dataview
> TABLE WITHOUT ID
> 	file.link as "Проект",
> 	status AS "Статус",
> 	priority AS "Приоритет",
> 	team AS "Команда"
> FROM "Projects"
> WHERE status != "done" AND status != "on-hold"
> SORT priority ASC
> ```

### 🎓 Houdini — Прогресс обучения

> [!example]- Детали прогресса (развернуть)
> ```dataview
> TABLE WITHOUT ID
>   file.link AS "Модуль",
>   length(filter(file.tasks, (t) => t.completed)) + " / " + length(file.tasks) AS "Прогресс"
> FROM "Houdini-Learning/Progress"
> SORT file.name ASC
> ```
---

## 🔥 Горящие задачи (Top 5)

```tasks
not done
priority is highest
priority is high
limit 5
sort by priority
```

---

## 🗺️ Навигация

> [!abstract] Карты и структуры
> *   [[Docs/Pathregist|🛣️ Pathregist (Роадмап)]]
> *   [[WTF.canvas|🗺️ Карта связей (Canvas)]]
> *   [[Inbox|📥 Inbox (Входящие)]]
> *   [[Tasks|📋 Глобальный список задач]]
