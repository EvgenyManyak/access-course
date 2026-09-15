-- =====================================================================================
-- crosstab-queries.sql — перекрёстные запросы TRANSFORM ... PIVOT. Курс «Microsoft Access
-- 2007–2016 для аналитика», модуль 8 «Аналитические запросы».
--
-- НАЗНАЧЕНИЕ ФАЙЛА
--   Четыре перекрёстных запроса к мини-датасету ShopData (Orders = 12, OrderDetails = 15,
--   Products = 8, Categories = 5). Перекрёстный запрос сворачивает длинный список в матрицу:
--   строки задаются в SELECT, столбцы — в PIVOT, содержимое ячеек — в TRANSFORM.
--
-- КАК ИСПОЛЬЗОВАТЬ
--   Вставьте ОДИН оператор целиком (от слова TRANSFORM до точки с запятой) в окно SQL-режима
--   конструктора запросов и выполните кнопкой «!». Строки «--» не копируйте: комментарии
--   в Access не работают — они только для чтения здесь. Тип запроса после вставки станет
--   «Перекрёстный» — так и должно быть.
--
-- ПОДСКАЗКА ПО СТРУКТУРЕ
--   TRANSFORM  <агрегат для ячеек>
--   SELECT     <поля строк> + итог по строке
--   FROM/WHERE <источник и фильтры>
--   GROUP BY   <поля строк>
--   PIVOT      <поле столбцов> [IN (фиксированный список)]
-- =====================================================================================


-- ---------------------------------------------------------------------------------------------
-- 1. Заказы по месяцам (строки) и статусам (столбцы), в ячейках COUNT
-- Столбцы создаются автоматически по фактическим значениям OrderStatus и сортируются по алфавиту:
--   canceled, delivered, processing, shipped.
-- Ожидаемый результат: 12 строк — по одной на месяц (в мини-датасете каждый месяц содержит
--   ровно один заказ), в каждой строке Итого = 1; сумма всех ячеек = 12.
--   Месяцы: 2017-03, 2017-04, 2017-06, 2017-08, 2017-10, 2017-11, 2018-01, 2018-02, 2018-03,
--   2018-04, 2018-05, 2018-06 (месяц 2017-05 и другие без заказов не появятся вовсе).
-- ---------------------------------------------------------------------------------------------
TRANSFORM COUNT(o.OrderID)
SELECT FORMAT(o.OrderDate, "yyyy-mm") AS Месяц, COUNT(o.OrderID) AS Итого
FROM Orders AS o
GROUP BY FORMAT(o.OrderDate, "yyyy-mm")
PIVOT o.OrderStatus;


-- ---------------------------------------------------------------------------------------------
-- 2. Выручка по месяцам (строки) и категориям (столбцы), в ячейках SUM
-- FORMAT(OrderDate,"yyyy-mm") даёт подписи строк вида 2017-03; отменённые заказы исключены.
-- Ожидаемый результат: 11 строк (месяцы 2017-03, 2017-04, 2017-06, 2017-08, 2017-10, 2018-01,
--   2018-02, 2018-03, 2018-04, 2018-05, 2018-06; месяца 2017-11 нет — там был только отменённый
--   O-10006). Столбцы по алфавиту: auto, bed_bath_table, computers_accessories, furniture_decor,
--   health_beauty. Контрольные итоги столбцов: auto 238,00; bed_bath_table 769,50;
--   computers_accessories 746,50; furniture_decor 640,00; health_beauty 631,50;
--   сумма всех ячеек и всех Итого = 3 025,50.
--   Пример строки 2017-03: bed_bath_table 189,90; health_beauty 179,80; Итого 369,70.
-- ---------------------------------------------------------------------------------------------
TRANSFORM SUM(od.Quantity * od.UnitPrice)
SELECT FORMAT(o.OrderDate, "yyyy-mm") AS Месяц, SUM(od.Quantity * od.UnitPrice) AS Итого
FROM ((Orders AS o INNER JOIN OrderDetails AS od ON o.OrderID = od.OrderID)
     INNER JOIN Products AS p ON od.ProductID = p.ProductID)
     INNER JOIN Categories AS cat ON p.CategoryID = cat.CategoryID
WHERE o.OrderStatus <> "canceled"
GROUP BY FORMAT(o.OrderDate, "yyyy-mm")
PIVOT cat.CategoryNameEn;


-- ---------------------------------------------------------------------------------------------
-- 3. Фиксированный список столбцов: PIVOT ... IN (...)
-- Список IN задаёт состав И порядок столбцов: только три категории, в заданном порядке,
-- даже если в данных есть другие. Фильтр WHERE по CategoryID (1, 2, 3) нужен, чтобы итог
-- столбца Итого совпадал с суммой трёх показанных столбцов: без него Итого включал бы
-- выручку furniture_decor и auto, которой среди столбцов нет.
-- Ожидаемый результат: 11 строк (те же месяцы, что в блоке 2); три столбца:
--   bed_bath_table, computers_accessories, health_beauty.
--   Контрольные итоги столбцов: bed_bath_table 769,50; computers_accessories 746,50;
--   health_beauty 631,50; сумма всех ячеек = 2 147,50.
-- ---------------------------------------------------------------------------------------------
TRANSFORM SUM(od.Quantity * od.UnitPrice)
SELECT FORMAT(o.OrderDate, "yyyy-mm") AS Месяц, SUM(od.Quantity * od.UnitPrice) AS Итого
FROM ((Orders AS o INNER JOIN OrderDetails AS od ON o.OrderID = od.OrderID)
     INNER JOIN Products AS p ON od.ProductID = p.ProductID)
     INNER JOIN Categories AS cat ON p.CategoryID = cat.CategoryID
WHERE o.OrderStatus <> "canceled" AND p.CategoryID IN (1, 2, 3)
GROUP BY FORMAT(o.OrderDate, "yyyy-mm")
PIVOT cat.CategoryNameEn IN ("bed_bath_table", "computers_accessories", "health_beauty");


-- ---------------------------------------------------------------------------------------------
-- 4. Перекрёстный запрос по кварталам
-- Код формата "yyyy-q" возвращает квартал числом 1–4, поэтому подписи строк вида 2017-1.
-- Ожидаемый результат: 6 строк:
--   2017-1: Итого 1 (delivered 1)             — заказ O-10001
--   2017-2: Итого 2 (delivered 2)             — O-10002, O-10003
--   2017-3: Итого 1 (delivered 1)             — O-10004
--   2017-4: Итого 2 (shipped 1, canceled 1)   — O-10005, O-10006
--   2018-1: Итого 3 (delivered 3)             — O-10007, O-10008, O-10009
--   2018-2: Итого 3 (delivered 2, processing 1) — O-10010, O-10011, O-10012
--   Сумма всех ячеек = 12.
--   В полной схеме кварталы можно взять из таблицы Calendar (CalYear, CalQuarter) соединением
--   по CalDate = OrderDate — результат совпадёт с этим запросом.
-- ---------------------------------------------------------------------------------------------
TRANSFORM COUNT(o.OrderID)
SELECT FORMAT(o.OrderDate, "yyyy-q") AS Квартал, COUNT(o.OrderID) AS Итого
FROM Orders AS o
GROUP BY FORMAT(o.OrderDate, "yyyy-q")
PIVOT o.OrderStatus;
