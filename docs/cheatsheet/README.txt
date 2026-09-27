Microsoft Access — SQL и функции: расширенный набор учебных карточек
=====================================================================

Расширение шпаргалки access-sql-cheatsheet.png: каждый логический блок
вынесен на отдельную страницу A4 (портрет) с крупным заголовком, теорией,
синтаксисом, примерами на одно и несколько значений, пояснением результата
и мини-иллюстрацией. Все примеры исходной шпаргалки сохранены и расширены.

Каждая карточка оператора снабжена собственной мини-иллюстрацией:
диаграммы-кружки для всех видов JOIN и логики AND/OR/NOT, воронки-фильтры
для условий, диалоговые окна Access для ошибок, календарь/часы для дат,
bar-диаграммы для агрегатов, сцены «до → после» для DML.

Состав (12 блоков, 1 блок = 1 страница PDF = 1 PNG):

  01_DDL.png         DDL — CREATE TABLE, ALTER TABLE, DROP TABLE, PK → FK
  02_DML.png         DML — INSERT INTO, UPDATE, DELETE + сцены «до → после»
  03_SELECT.png      SELECT — поля, DISTINCT, AS, ORDER BY, TOP
  04_CONDITIONS.png  Условия отбора — =, IN, BETWEEN, LIKE, IS NULL, AND, OR, NOT
  05_JOIN.png        JOIN — INNER, LEFT, RIGHT, CROSS, SELF + FULL OUTER через UNION
  06_CONVERSION.png  Преобразование типов — Nz, IsNull, Val, CCur, CInt/CLng, CDate, Replace, IIf
  07_TEXT.png        Текст — &, Left, Right, Mid, Len, UCase/LCase, Trim, InStr, Replace
  08_DATES.png       Даты — Date, Now, Year/Month/Day, DateAdd, DateDiff, Format
  09_AGGREGATES.png  Агрегаты — Count, Sum, Avg, Min, Max, GROUP BY, HAVING
  10_VALIDATION.png  Проверка значений поля — Короткий текст, размер 40, In (...)
  11_DATA_TYPES.png  Типы данных Access — 8 типов с иконками и примерами полей
  12_ERRORS.png      Частые ошибки — ошибка → причина → исправление

Файлы:
  access-sql-cheatsheet-v2.pdf — многостраничный векторный PDF (12 страниц A4,
                                 текст выделяется и ищется)
  01..12 *.png                 — каждая страница отдельно, 2480×3508 px, 300 dpi
  html/                        — исходники страниц (HTML/CSS) и скрипты сборки
                                 (render.py — PNG, build_combined.py + all.html — PDF)

Формат: A4 вертикально, 300 dpi, шрифты Carlito / DejaVu Sans Mono,
стиль и палитра исходной шпаргалки access-course.

Совместимость: Microsoft Access 2007–2016, русская локаль
(в выражениях аргументы разделяются точкой с запятой).
