Attribute VB_Name = "modImportExamples"
' =====================================================================================
' modImportExamples — импорт данных в Access (файлы к курсу
' «Microsoft Access 2007–2016 для аналитика», модули 12–14).
'
' НАЗНАЧЕНИЕ
'   Пять готовых процедур импорта и обслуживания связанных таблиц:
'     btnImportCsv_Click  — кнопка «Импорт клиентов»: CSV в Customers_Staging,
'                           затем перенос новых клиентов запросом на добавление
'     ImportExcelSheet    — лист Excel в Products_Staging (TransferSpreadsheet)
'     RelinkTable         — переподключение ОДНОЙ связанной таблицы к новому файлу БД
'     RefreshAllLinks     — переподключение ВСЕХ связанных таблиц (база «переехала»)
'     ImportFolderCsvs    — импорт всех CSV-файлов из папки import одним запуском
'
' КАК ПОДКЛЮЧИТЬ
'   VBE (Alt+F11) → File → Import File... → выберите этот файл .bas;
'   либо Insert → Module (обычный модуль) и вставьте текст целиком.
'   Модуль появится в Project Explorer (Ctrl+R) в папке Modules под именем modImportExamples.
'   Процедура btnImportCsv_Click рассчитана на кнопку: перенесите её текст в модуль
'   формы (как в модуле 12) — код событий кнопок Access ищет в модуле формы.
'
' КАК ЗАПУСТИТЬ
'   Поставьте курсор внутрь нужной процедуры и нажмите F5 (Run → Run Sub/UserForm);
'   либо пошагово — F8; либо кнопкой на форме (свойство «Нажатие кнопки» →
'   [Процедура обработки событий]).
'
' ТРЕБОВАНИЯ
'   Файлы лежат в папке базы (CurrentProject.Path):
'     customers_import.csv  — еженедельная выгрузка клиентов (3 строки C-0009…C-0011);
'     products_import.xlsx  — прайс-лист в Excel (3 строки P-6001…P-6003);
'     подпапка import\      — CSV одной структуры для ImportFolderCsvs.
'   Сохранённые объекты:
'     таблицы Customers_Staging и Products_Staging (модули 4 и 12);
'     спецификация импорта spec_ИмпортКлиентов (имя произвольное — см. процедуру 1);
'     запрос на добавление qry_ПереносКлиентов — переносит новых клиентов из staging
'     в Customers (в модуле 4 такой запрос сохранялся как qry_ПереносКлиентов).
'   В References должна стоять библиотека Microsoft Office XX.0 Access database engine
'   Object Library (DAO) — она включена по умолчанию.
'   При открытии базы нажмите «Включить содержимое» на жёлтой панели безопасности:
'   без этого VBA отключён, F5 и кнопки не работают.
' =====================================================================================

Option Compare Database
Option Explicit

' ---------------------------------------------------------------------------------------------
' 1. Кнопка «Импорт клиентов»: CSV → Customers_Staging → перенос новых в Customers
' ---------------------------------------------------------------------------------------------
Public Sub btnImportCsv_Click()
    ' Импортирует еженедельный файл customers_import.csv в staging-таблицу Customers_Staging,
    ' затем открывает запрос на добавление qry_ПереносКлиентов — он переносит в Customers
    ' только тех клиентов, которых ещё нет (повторы отсекаются условием в самом запросе).
    ' Ожидаемый результат: в Customers_Staging появились строки C-0009…C-0011 (3 строки),
    ' ZipCodePrefix = 04578 сохранил ведущий ноль, запрос добавил 3 новых клиента.

    Const SPEC_NAME As String = "spec_ИмпортКлиентов"  ' имя сохранённой спецификации импорта

    Dim strFile As String                 ' полный путь к CSV-файлу

    On Error GoTo ErrHandler              ' при ошибке исполнения — переход на метку ErrHandler

    strFile = CurrentProject.Path & "\customers_import.csv"   ' файл лежит в папке базы

    If Len(SPEC_NAME) > 0 Then
        ' Вариант А: со спецификацией — разделитель, кодировка и типы полей заданы заранее.
        ' Имя спецификации произвольное, главное — совпадение с сохранённой в вашей базе:
        ' в модуле 4 она называлась spec_ИмпортКлиентов, в модуле 12 — «ИмпортКлиентов».
        DoCmd.TransferText acImportDelim, SPEC_NAME, "Customers_Staging", strFile, True
    Else
        ' Вариант Б: без спецификации (пустой второй аргумент) — настройки Access по умолчанию;
        ' тогда ZipCodePrefix может определиться числом и потерять ведущий ноль 01310 → 1310.
        DoCmd.TransferText acImportDelim, , "Customers_Staging", strFile, True
    End If
    ' Разбор аргументов TransferText:
    '   acImportDelim       — импорт текста с разделителями (не экспорт и не связывание);
    '   SPEC_NAME           — сохранённая спецификация импорта (или пусто — вариант Б);
    '   "Customers_Staging" — таблица-приёмник (staging, а не боевая таблица Customers!);
    '   strFile             — полный путь к файлу;
    '   True                — первая строка файла содержит имена полей.

    DoCmd.OpenQuery "qry_ПереносКлиентов"
    ' Запрос на добавление: Access спросит «Вы собираетесь добавить N строк» — подтвердите.
    ' Имя запроса должно совпадать с вашим сохранённым запросом из модуля 4
    ' (там он сохранялся как qry_ПереносКлиентов — подставьте своё имя, если оно другое).

    MsgBox "Импорт завершён: customers_import.csv загружен в Customers_Staging," & _
           " новые клиенты перенесены в Customers.", vbInformation, "Импорт клиентов"

    Exit Sub                              ' штатный выход — ОБЯЗАТЕЛЬНО выше обработчика

ErrHandler:                               ' обработчик ошибок
    ' Частые причины: файла нет или он занят другой программой (3051/3625),
    ' нет таблицы Customers_Staging (3078), нет спецификации с таким именем (3625).
    MsgBox "Не удалось импортировать customers_import.csv." & vbCrLf & _
           "Проверьте, что файл лежит в папке базы: " & CurrentProject.Path & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Импорт клиентов"
End Sub

' ---------------------------------------------------------------------------------------------
' 2. Импорт листа Excel в Products_Staging
' ---------------------------------------------------------------------------------------------
Public Sub ImportExcelSheet()
    ' Импортирует лист книги products_import.xlsx в таблицу Products_Staging.
    ' Столбцы листа: ProductID, ProductName, CategoryID, UnitPrice.
    ' В файле — 3 новых товара: P-6001 «Тент автомобильный» (категория 5),
    ' P-6002 «GPS-навигатор» (категория 3), P-6003 «Одеяло туристическое» (категория 1).
    ' Ожидаемый результат: в Products_Staging добавились 3 строки. Повторный запуск
    ' задвоит их — staging принимает всё, от повторов защищает запрос переноса.

    Dim strFile As String                 ' полный путь к книге Excel

    On Error GoTo ErrHandler

    strFile = CurrentProject.Path & "\products_import.xlsx"

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Products_Staging", _
                              strFile, True
    ' Разбор аргументов TransferSpreadsheet:
    '   acImport                      — режим импорта (acLink связал бы, а не скопировал);
    '   acSpreadsheetTypeExcel12Xml   — формат .xlsx (для старых .xls — acSpreadsheetTypeExcel8);
    '   "Products_Staging"            — таблица-приёмник (Access создаст её, если ещё нет);
    '   strFile                       — полный путь к книге Excel;
    '   True                          — первая строка листа содержит заголовки полей.

    MsgBox "Импорт завершён: products_import.xlsx загружен в Products_Staging.", _
           vbInformation, "Импорт Excel"

    Exit Sub

ErrHandler:
    ' Частые причины: книга открыта в Excel (3051), не тот путь или имя файла.
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Импорт Excel"
End Sub

' ---------------------------------------------------------------------------------------------
' 3. Переподключение одной связанной таблицы к новому файлу БД
' ---------------------------------------------------------------------------------------------
Private Function RelinkTable(ByRef tdf As DAO.TableDef, ByVal strBEPath As String) As Boolean
    ' Меняет строку подключения связанной таблицы на новый путь и обновляет связь.
    ' Параметры:
    '   tdf        — TableDef связанной таблицы; передаётся ByRef: функция меняет
    '                сам объект базы, а не его копию;
    '   strBEPath  — полный путь к новому файлу данных (например,
    '                C:\ShopData\shopdata_be.accdb).
    ' Возвращает True — связь обновлена; False — произошла ошибка (например,
    ' файл данных не найден по указанному пути). Решение принимает вызывающий код.

    On Error GoTo ErrHandler

    tdf.Connect = ";DATABASE=" & strBEPath    ' строка подключения таблиц Access: ;DATABASE=путь
    tdf.RefreshLink                           ' немедленно применяем новую строку подключения

    RelinkTable = True                        ' успех
    Exit Function

ErrHandler:
    ' Частая причина: файла БД нет по этому пути (3044) или он занят монопольно.
    RelinkTable = False                       ' неудача — без аварийного останова
End Function

' ---------------------------------------------------------------------------------------------
' 4. Переподключение всех связанных таблиц (сценарий «файл данных переехал»)
' ---------------------------------------------------------------------------------------------
Public Sub RefreshAllLinks(CurrentBackendPath As String)
    ' После разделения базы (модуль 14) таблицы живут в файле данных shopdata_be.accdb,
    ' а формы, запросы и отчёты — в клиентском файле. Если файл данных переехал на другой
    ' диск или сетевой ресурс, все связи нужно переписать на новый путь.
    ' Запуск: в окне Immediate (Ctrl+G) наберите и нажмите Enter:
    '   RefreshAllLinks "C:\ShopData\shopdata_be.accdb"
    ' либо повесьте процедуру на кнопку служебной формы.
    ' Внимание: связи с листами Excel и текстовыми файлами этот код не переподключит
    ' (у них строка подключения другого вида — они попадут в счётчик ошибок);
    ' для них используйте Диспетчер связанных таблиц.

    Dim db As DAO.Database                ' текущая (клиентская) база
    Dim tdf As DAO.TableDef               ' очередной TableDef в цикле
    Dim lngOK As Long                     ' сколько таблиц переподключено успешно
    Dim lngFail As Long                   ' сколько таблиц переподключить не удалось

    On Error GoTo ErrHandler

    If Dir(CurrentBackendPath) = "" Then  ' файла данных по пути нет — идти дальше бессмысленно
        MsgBox "Файл данных не найден: " & CurrentBackendPath, _
               vbExclamation, "Обновление связей"
        Exit Sub                          ' ранний выход — тоже выше обработчика
    End If

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        ' У локальных таблиц (Customers, Orders и другие) свойство Connect пустое —
        ' их пропускаем; у связанных оно заполнено, например:
        ' ";DATABASE=C:\ShopData\shopdata_be.accdb".
        If Len(tdf.Connect) > 0 Then
            If RelinkTable(tdf, CurrentBackendPath) Then
                lngOK = lngOK + 1         ' связь обновлена
            Else
                lngFail = lngFail + 1     ' файл занят, путь неверен и прочее
            End If
        End If
    Next tdf

    MsgBox "Переподключено таблиц: " & lngOK & ", с ошибками: " & lngFail & "." & vbCrLf & _
           "Файл данных: " & CurrentBackendPath, _
           IIf(lngFail = 0, vbInformation, vbExclamation), "Обновление связей"

    Set db = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Обновление связей"
End Sub

' ---------------------------------------------------------------------------------------------
' 5. Импорт всех CSV из папки import одним запуском
' ---------------------------------------------------------------------------------------------
Public Sub ImportFolderCsvs()
    ' Импортирует ВСЕ файлы *.csv из подпапки import (рядом с базой) в Customers_Staging.
    ' Сценарий: маркетплейс кладёт в папку еженедельные выгрузки клиентов одной структуры —
    ' staging накапливает строки из всех файлов, а запрос на добавление (процедура 1)
    ' переносит в Customers только новых. Ожидаемый результат: счётчик равен числу
    ' CSV-файлов в папке (например, customers_2018-07-02.csv, customers_2018-07-09.csv,
    ' customers_2018-07-16.csv — в MsgBox будет 3).

    Const SPEC_NAME As String = "spec_ИмпортКлиентов"  ' сохранённая спецификация импорта

    Dim strFolder As String               ' папка с файлами (обязательно со слешем на конце)
    Dim strFile As String                 ' имя очередного файла (без пути)
    Dim lngFiles As Long                  ' счётчик импортированных файлов

    On Error GoTo ErrHandler

    strFolder = CurrentProject.Path & "\import\"

    If Dir(strFolder, vbDirectory) = "" Then          ' папки нет — предупреждение и выход
        MsgBox "Папка не найдена: " & strFolder & "." & vbCrLf & _
               "Создайте подпапку import рядом с базой и положите туда CSV-файлы.", _
               vbExclamation, "Импорт папки CSV"
        Exit Sub
    End If

    strFile = Dir(strFolder & "*.csv")                ' имя ПЕРВОГО CSV-файла (Dir с аргументом)

    Do While Len(strFile) > 0                         ' пока Dir возвращает непустое имя
        DoCmd.TransferText acImportDelim, SPEC_NAME, "Customers_Staging", _
                           strFolder & strFile, True   ' путь к файлу = папка + имя
        lngFiles = lngFiles + 1                       ' файл отправлен на импорт
        strFile = Dir()                               ' следующий файл: Dir БЕЗ аргументов
    Loop                                              ' продолжает начатый обход папки

    MsgBox "Импортировано файлов: " & lngFiles & " из папки " & strFolder, _
           vbInformation, "Импорт папки CSV"

    Exit Sub

ErrHandler:
    ' Частая причина: один из файлов занят другой программой (3051) или структура файла
    ' не совпадает со спецификацией — импорт прерывается на текущем файле.
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Импорт папки CSV"
End Sub
