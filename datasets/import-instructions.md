# Инструкция по импорту данных (import-instructions)

Файл описывает три пути получения учебных данных курса ShopData. Выберите один: **A — основной датасет Olist**, **B — альтернативный датасет**, **C — мини-датасет без Kaggle**. После любого пути у вас должны работать все задания модулей 02–15.

Общее правило курса: **сначала импорт во staging-таблицы, потом проверка и очистка, потом перенос в основные таблицы, и только потом — связи**. Пошагово это разобрано в модулях 04 и 05.

## Подготовка рабочей папки

1. Создайте папку проекта, например `C:\Courses\ShopData\`.
2. Внутри создайте: `import\` (сюда распаковываются CSV), `base\` (здесь будут `.accdb`), `export\` (сюда пишутся выгрузки).
3. Создайте в Access пустую базу `base\ShopData.accdb` (Файл → Создать → Пустая база данных рабочего стола).

## Путь A — основной датасет Olist (рекомендуется)

Источник: [kaggle.com/datasets/olistbr/brazilian-ecommerce](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce) (нужна бесплатная учётная запись Kaggle; лицензия CC BY-NC-SA 4.0).

1. Нажмите **Download** на странице набора и распакуйте архив в папку `import\`. Появятся 9 CSV-файлов.
2. Для курса используются 5 файлов: `olist_customers_dataset.csv`, `olist_orders_dataset.csv`, `olist_order_items_dataset.csv`, `olist_products_dataset.csv`, `product_category_name_translation.csv`.
3. Импорт каждого файла — по инструкции модуля 04 (Внешние данные → Новый источник данных → Из файла → Тестовый файл). Разделитель — запятая, первая строка — имена полей, кодировка UTF-8.
4. Маппинг файлов на таблицы курса:

| Файл Olist | Staging-таблица | Основная таблица курса | Правила преобразования |
|---|---|---|---|
| `olist_customers_dataset.csv` | `Customers_Staging` | `Customers` | `customer_id` → CustomerID (текст 32); `customer_zip_code_prefix` → ZipCodePrefix **тип «Короткий текст»**, иначе теряется ведущий ноль; `customer_city`, `customer_state` → City, State; CustomerName и Segment — учебные поля, заполняются из учебного пакета или запросом-генератором |
| `olist_orders_dataset.csv` | `Orders_Staging` | `Orders` | `order_id` → OrderID; `customer_id` → CustomerID; `order_status` → OrderStatus; `order_purchase_timestamp` → OrderDate; `order_delivered_customer_date` → DeliveredDate (пусто = NULL, это нормально); `order_estimated_delivery_date` → EstimatedDate; ManagerID — учебное поле (см. шаг 5) |
| `olist_order_items_dataset.csv` | `OrderDetails_Staging` | `OrderDetails` | `order_id` → OrderID; `product_id` → ProductID; `price` → UnitPrice (Денежный); `freight_value` → FreightValue; `order_item_id` не является уникальным ключом — суррогатный OrderDetailID делает Счётчик; Quantity по умолчанию 1 |
| `olist_products_dataset.csv` | `Products_Staging` | `Products` | `product_id` → ProductID; `product_category_name` → связка на Categories; `product_weight_g` → WeightG; габариты → LengthCm/HeightCm/WidthCm; ProductName, UnitPrice — учебные поля |
| `product_category_name_translation.csv` | `Categories_Staging` | `Categories` | `product_category_name` → CategoryName; `product_category_name_english` → CategoryNameEn; после импорта CategoryID (Счётчик) заполнится сам |

5. Учебные поля (ManagerID, EmployeeID-привязка): заполните запросом на обновление, который случайно-детерминированно распределяет менеджеров: `UPDATE Orders SET ManagerID = 1 + (Asc(Left([OrderID],1)) Mod 6);` — либо вручную для первых 12 заказов. Таблица `Employees` заполняется из учебного пакета (6 строк) или по образцу из `datasets/data-dictionary.md`.
6. Таблица `Calendar`: создайте через генератор (VBA-процедура из `vba/export-examples.bas`-стиля или цикл INSERT) на диапазон 2016-10-01 … 2018-10-31, либо импортируйте готовый CSV-календарь. Поля — в data-dictionary.md.
7. Очистка (модуль 05): дубликаты покупателей, пропуски дат доставки, проверка `UnitPrice >= 0`.
8. Связи (модуль 03): создайте 6 связей из data-dictionary.md; для Orders → OrderDetails включите каскадное удаление.

⚠️ Объём Olist — 100 тыс. заказов: упражнения по формам и отчётам удобнее делать на выборке (запрос make-table на 12–20 заказов в учебную копию таблиц), а полный объём оставить для агрегатов.

## Путь B — альтернативный датасет

**Superstore** ([kaggle.com/datasets/vivek468/superstore-dataset-final](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final), 1 CSV ≈ 10 тыс. строк):

| Колонка Superstore | Таблица курса | Поле |
|---|---|---|
| Customer ID / Customer Name / Segment / City / State / Postal Code / Region | Customers | CustomerID / CustomerName / Segment / City / State / ZipCodePrefix / — |
| Product ID / Product Name / Category / Sub-Category | Products, Categories | ProductID / ProductName / CategoryID (через Categories) |
| Order ID / Order Date / Ship Date / Ship Mode | Orders | OrderID / OrderDate / DeliveredDate (Ship Date) / ShipMode → можно добавить поле |
| Sales / Quantity / Discount | OrderDetails | UnitPrice (= Sales/Quantity), Quantity, Discount (добавить поле) |
| Row ID | — | не переносится (суррогатный Счётчик) |

Employees и Calendar — синтетические, как в пути A, шаги 5–6. Эталонные суммы курса на Superstore не воспроизводятся — контрольные числа из data-dictionary.md применять нельзя.

**Online Retail (UCI)** ([kaggle.com/datasets/carrie1/ecommerce-data](https://www.kaggle.com/datasets/carrie1/ecommerce-data)): перекодируйте CSV в UTF-8 (например, в Excel «Сохранить как → CSV UTF-8»); InvoiceNo → Orders (учтите возвраты с буквой C и отрицательной Quantity — материал для очистки, модуль 05); StockCode → Products; CustomerID есть не у всех строк — отличное упражнение на пропуски.

## Путь C — мини-датасет без Kaggle

1. Откройте `ShopData.accdb`.
2. Выполните SQL-файлы из каталога `sql/` по порядку: сначала `CREATE TABLE`-операторы (есть в начале каждого файла), затем `INSERT INTO`. Каждое окно SQL в Access выполняет **один оператор** — копируйте блоками по одному.
3. Проще всего начать с запросов из `sql/select-examples.sql`: шапка файла объясняет правила.
4. В результате получите те же 7 таблиц, что описаны в `data-dictionary.md`, с контрольными числами: 12 заказов, 15 позиций, выручка 3 025,50 (без отменённых).

## Типовые проблемы импорта и решения

| Симптом | Причина | Решение |
|---|---|---|
| ZipCodePrefix 01310 стал 1310 | Поле импортировано как Числовой | Удалить таблицу; повторить импорт, тип поля — Короткий текст (модуль 04) |
| «Ключевая ошибка» при импорте | Дубликаты ключей в CSV | Импорт в staging → дедупликация (модуль 05) → перенос |
| Даты — текстом | Нестандартный формат | Импорт как текст → запрос на обновление с CDate → тип Дата/время |
| Кракозябры вместо букв | Кодировка не UTF-8 | Перекодировать файл или выбрать кодировку в расширенных параметрах мастера |
| Число 189,9 не распознаётся | Десятичная запятая при русской локали | Расширенные параметры мастера: десятичный разделитель «запятая» |

## Что проверить после импорта (чек-лист)

- [ ] В каждой таблице количество строк соответствует ожиданию (сверьте с мастером импорта);
- [ ] ZipCodePrefix — текст, ведущие нули на месте;
- [ ] Даты — тип Дата/время, сортировка по дате работает;
- [ ] Денежные поля — тип Денежный, отрицательных цен нет;
- [ ] В OrderDetails нет позиций с несуществующим OrderID/ProductID (запрос несоответствий, модуль 05);
- [ ] Связи созданы без ошибок, схема данных выглядит как на рисунке модуля 03.
