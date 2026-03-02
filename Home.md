# 🧠 Equilibrium

> «Единая система знаний — основа инженерного мастерства»
> — В. & А.

---

## 🚀 Быстрые действия

```button
name ➕ Новая идея
type command
action Templater: Create new note from template
```

```button
name 📝 Новый проект
type command
action Templater: Create new note from template
```

```button
name ✅ Задачи
type link
action obsidian://open?file=Tasks
```

```button
name ♻️ Синхронизировать
type command
action Git: Commit-and-sync
```

---

## 📊 Проекты

```dataview
TABLE status AS "Статус", priority AS "Приоритет", team AS "Ответственные"
FROM "Projects"
WHERE priority AND status
SORT priority ASC
```

---

## 📚 Houdini — Прогресс

```dataview
TABLE WITHOUT ID
  file.link AS "Участник",
  length(filter(file.tasks, (t) => t.completed)) AS "✅ Выполнено",
  length(filter(file.tasks, (t) => !t.completed)) AS "⬜ Осталось"
FROM "Houdini-Learning/Progress"
SORT file.name ASC
```

---

## 🔴 Активные задачи

```tasks
not done
priority is highest
OR priority is high
limit 5
sort by priority
```

---

## 🗺️ Карта проекта

[[WTF.canvas]]