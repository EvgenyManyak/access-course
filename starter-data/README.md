# Стартовый комплект данных (starter-data)

Комплект позволяет начать обучение **без Kaggle, без регистрации и без сборки данных вручную**: шесть готовых CSV-файлов с эталонным мини-датасетом ShopData импортируются в Access за 15–20 минут, и все контрольные числа курса (модули 01–15) сходятся с первого запроса.

## Файлы комплекта

| Файл | Таблица Access | Строк | Первичный ключ | Назначение |
|---|---|---|---|---|
| `categories.csv` | Categories | 5 | CategoryID | Справочник категорий товаров (португальское и английское имя) |
| `products.csv` | Products | 8 | ProductID | Товары: название, категория, цена, вес, габариты |
| `customers.csv` | Customers | 8 | CustomerID | Покупатели: имя, почтовый префикс, город, штат, сегмент |
| `customers_archive.csv` | Customers | 11 | CustomerID | Покупатели: имя, почтовый префикс, город, штат, сегмент. Содержит дубликаты записей |
| `employees.csv` | Employees | 6 | EmployeeID | Сотрудники — менеджеры заказов: роль, город, дата найма |
| `orders.csv` | Orders | 12 | OrderID | Заказы: покупатель, менеджер, статус, три даты |
| `order-details.csv` | OrderDetails | 15 | OrderDetailID | Позиции заказов: товар, количество, цена, стоимость доставки |
| `README.md` | — | — | — | Этот файл: состав, количества, образец данных |
| `data-dictionary.md` | — | — | — | Словарь: все поля, типы Access, ключи, связи, назначение |
| `import-checklist.md` | — | — | — | Пошаговый чек-лист импорта и решения типичных ошибок |

Седьмая таблица курса — служебный календарь **Calendar** (761 дата, 01.10.2016–31.10.2018) — не входит в стартовый комплект: она нужна начиная с модуля 04 и создаётся готовым VBA-генератором (`../vba/calendar-generator.bas`) или импортом файла `../datasets/calendar.csv`. Инструкция — в `../datasets/import-instructions.md`, раздел «Таблица Calendar — три способа создания».

## Формат файлов

- Кодировка **UTF-8 с BOM**, переводы строк CRLF — как в `datasets/calendar.csv`;
- разделитель — **запятая**, ограничитель текста — двойная кавычка (не используется, но включите в мастере);
- даты — **ISO (ГМД)**: `2017-03-15`; в мастере импорта откройте «Дополнительно…» и выберите порядок дат **ГМД**;
- десятичный разделитель цен — **точка**: `189.90` (на русской Windows поле может импортироваться текстом — решение в import-checklist.md, ошибка № 4);
- пустые поля (например, `DeliveredDate` у недоставленных заказов) — пусто = NULL, это нормально.

## Ожидаемое количество записей после импорта

| Показатель | Значение |
|---|---|
| Categories / Products / Customers / Employees / Orders / OrderDetails | **5 / 8 / 8 / 6 / 12 / 15** |
| Статусы заказов | 9 delivered, 1 shipped (O-10005), 1 canceled (O-10006), 1 processing (O-10010) |
| Период заказов | 2017-03-15 … 2018-06-12 (каждый месяц — не более одного заказа) |
| Выручка `SUM(UnitPrice*Quantity)`, все статусы | 3 115,40 |
| Выручка без отменённых | **3 025,50** |
| Средний срок доставки (delivered) | 102 дня / 9 заказов ≈ 11,3 дня |

Объём подобран сознательно компактным (5–15 строк в таблице): каждая строка задействована в заданиях, а все суммы и списки результатов проверяются глазами построчно. Для практики на большом объёме используется полный датасет Olist (путь А в `../datasets/import-instructions.md`).

## Образец корректно импортированных данных

Так должна выглядеть таблица Customers после импорта (режим таблицы; на русской локали Access покажет поля слева направо в порядке полей файла):

| CustomerID | CustomerName | ZipCodePrefix | City | State | Segment |
|---|---|---|---|---|---|
| C-0001 | Ana Ribeiro | 01310 | sao paulo | SP | Retail |
| C-0002 | Bruno Carvalho | 20040 | rio de janeiro | RJ | VIP |
| C-0003 | Carla Mendes | 30120 | belo horizonte | MG | Wholesale |
| C-0004 | Diego Silva | 80410 | curitiba | PR | Retail |
| C-0005 | Elena Costa | 50050 | recife | PE | Retail |
| C-0006 | Felipe Alves | 90010 | porto alegre | RS | Wholesale |
| C-0007 | Gabriela Lima | 40020 | salvador | BA | VIP |
| C-0008 | Henrique Souza | 69050 | manaus | AM | Retail |

Первые строки остальных таблиц:

- **categories.csv**: `1, cama_mesa_banho, bed_bath_table` · `2, beleza_saude, health_beauty` … `5, automotivo, auto`;
- **products.csv**: `P-1001, Набор постельного белья 2-сп, 1, 189.90, 900, …`;
- **employees.csv**: `1, Ana Souza, S-0001, Менеджер, sao paulo, SP, 2015-03-02`;
- **orders.csv**: `O-10001, C-0001, 2, delivered, 2017-03-15, 2017-03-24, 2017-03-28`;
- **order-details.csv**: `1, O-10001, P-1001, 1, 189.90, 22.50, 2017-03-20`.

## Связи между таблицами

Все пять связей строятся в «Схеме данных» после импорта (подробно — модуль 03 и import-checklist.md):

| Связь | Поля | Тип |
|---|---|---|
| Categories → Products | CategoryID → CategoryID | 1:М |
| Customers → Orders | CustomerID → CustomerID | 1:М |
| Employees → Orders | EmployeeID → ManagerID | 1:М |
| Orders → OrderDetails | OrderID → OrderID | 1:М (+ каскадное удаление) |
| Products → OrderDetails | ProductID → ProductID | 1:М |

Шестая связь Calendar → Orders (по CalDate = OrderDate) добавляется в модуле 04 после создания календаря.

## С чего начать

1. Пройдите страницу **«Начало работы»** (`../getting-started.html`) — она ведёт от установки Access до первого запроса по шагам.
2. Импортируйте файлы по чек-листу `import-checklist.md` (порядок: categories → products → customers → employees → orders → order-details).
3. Сверьте количества строк и первый запрос `SELECT TOP 5` — эталоны в import-checklist.md и модуле 01.
