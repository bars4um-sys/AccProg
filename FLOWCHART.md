# AccProg — Блок-схемы

Диаграммы в формате **Mermaid** (отображаются на GitHub, в VS Code с расширением Mermaid,
в Obsidian и на https://mermaid.live).

---

## 1. Структура интерфейса (4 панели)

```mermaid
flowchart TB
    Header["ШАПКА — AccProg"]

    subgraph App["Оболочка приложения (shell)"]
        direction LR
        Sidebar["ЛЕВАЯ ПАНЕЛЬ<br/>Выбор функции<br/>(меню из реестра)"]
        subgraph Right["Рабочая область"]
            direction TB
            P1["ПАНЕЛЬ 1 — Ввод данных<br/>(загрузка таблиц / параметры)"]
            P2["ПАНЕЛЬ 2 — Промежуточные операции<br/>(кнопки шагов)"]
            P3["ПАНЕЛЬ 3 — Результат<br/>(вывод + кнопка «Скачать»)"]
            P1 --> P2 --> P3
        end
        Sidebar -->|активирует функцию| Right
    end

    Header --> App
```

---

## 2. Сценарий работы пользователя (жизненный цикл функции)

```mermaid
flowchart TD
    Start([Пользователь открыл AccProg]) --> Load[Ядро core.js читает реестр функций]
    Load --> Menu[Левая панель: список функций]
    Menu --> Pick{Выбрана функция?}
    Pick -- нет --> Menu
    Pick -- да --> Activate["shell.activate(id):<br/>очистка панелей 1-2-3"]

    Activate --> RenderInput["fn.renderInput(ctx)<br/>-> Панель 1 (загрузка данных)"]
    Activate --> RenderActions["fn.actions[]<br/>-> Панель 2 (кнопки)"]

    RenderInput --> Upload[Пользователь загружает таблицы]
    Upload --> ReadFile["utils.readFile + XLSX.read<br/>нормализация данных"]
    ReadFile --> StateOK{Данные валидны?}
    StateOK -- нет --> Err[Показать ошибку в Панели 1]
    Err --> Upload
    StateOK -- да --> Ready[Кнопки в Панели 2 активны]

    Ready --> Click[Нажата кнопка операции]
    Click --> Run["action.run(ctx)<br/>(сравнение / дубли / агрегация)"]
    Run --> RenderOut["fn.renderOutput(ctx, result)<br/>-> Панель 3 (таблица/список)"]
    RenderOut --> Download{Скачать результат?}
    Download -- да --> Export["fn.exportResult(ctx, result)<br/>XLSX.writeFile -> файл"]
    Download -- нет --> Done([Готово])
    Export --> Done
```

---

## 3. Архитектура файлов и поток подключения

```mermaid
flowchart LR
    subgraph Shell["Оболочка"]
        IDX["index.html"]
        CSS["assets/css/app.css"]
        CORE["assets/js/core.js<br/>(AccProg.register, роутинг)"]
        UTILS["assets/js/utils.js<br/>(Excel/CSV/нормализация/экспорт)"]
        VENDOR["assets/vendor/xlsx.full.min.js"]
    end

    subgraph Registry["Реестр функций"]
        REG["functions.yml / functions.json"]
    end

    subgraph Functions["Модули-функции"]
        F1["compare-tables.js<br/>(из SearchDiff)"]
        F2["advance-vs-bank.js<br/>(из SearchFall)"]
        FN["...новые функции"]
    end

    IDX --> CSS
    IDX --> VENDOR
    IDX --> UTILS
    IDX --> CORE
    CORE --> REG
    REG --> F1
    REG --> F2
    REG --> FN
    F1 --> UTILS
    F2 --> UTILS
    FN --> UTILS
```

---

## 4. Контракт модуля и взаимодействие с оболочкой

```mermaid
sequenceDiagram
    participant U as Пользователь
    participant S as Оболочка (core.js)
    participant F as Модуль-функция
    participant X as utils.js / XLSX

    U->>S: Выбор функции в меню
    S->>F: renderInput(ctx)
    S->>F: получить actions[]
    F-->>S: список кнопок -> Панель 2
    U->>S: Загрузка файлов
    S->>X: readFile() + normalize()
    X-->>S: данные в ctx.state
    U->>S: Нажатие кнопки
    S->>F: action.run(ctx)
    F->>X: расчёт (сравнение/дубли)
    F->>S: renderOutput(ctx, result) -> Панель 3
    U->>S: «Скачать результат»
    S->>F: exportResult(ctx, result)
    F->>X: XLSX.writeFile()
    X-->>U: файл .xlsx
```

---

## 5. Механизм расширения (как добавить функцию)

```mermaid
flowchart TD
    A([Нужна новая функция]) --> B[Создать assets/functions/&lt;id&gt;.js<br/>по контракту: renderInput / actions / renderOutput / exportResult]
    B --> C[Добавить запись в реестр<br/>functions.yml / functions.json]
    C --> D{Используется Jekyll?}
    D -- да --> E[Добавить _functions/&lt;id&gt;.md с описанием]
    D -- нет --> F[Ничего больше не нужно]
    E --> G([Функция появилась в меню])
    F --> G
    G --> H[Остальной код НЕ изменялся ✔]
```
