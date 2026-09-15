Attribute VB_Name = "modExportExamples"
' =====================================================================================
' modExportExamples — экспорт данных из Access (файлы к курсу
' «Microsoft Access 2007–2016 для аналитика», модули 12 и 13).
'
' НАЗНАЧЕНИЕ
'   Четыре готовые процедуры экспорта:
'     ExportCategoryToXlsx  — запрос в Excel (DoCmd.OutputTo, acFormatXLSX)
'     ExportReportToPdf     — отчёт в PDF (DoCmd.OutputTo, acFormatPDF)
'     ExportQueryToCsv      — запрос в CSV (DoCmd.TransferText, со спецификацией и без)
'     ExportAllMonthly      — цикл Do While по месяцам: файл CSV на каждый месяц
'
' КАК ПОДКЛЮЧИТЬ
'   VBE (Alt+F11) → File → Import File... → выберите этот файл .bas;
'   либо Insert → Module (обычный модуль) и вставьте текст целиком.
'   Модуль появится в окне Project Explorer (Ctrl+R) в папке Modules под именем modExportExamples.
'
' КАК ЗАПУСТИТЬ
'   Поставьте курсор внутрь нужной процедуры и нажмите F5 (Run → Run Sub/UserForm);
'   либо пошагово — F8; либо повесьте процедуру на кнопку формы (свойство «Нажатие кнопки»
'   → [Процедура обработки событий] → ExportCategoryToXlsx).
'
' ТРЕБОВАНИЯ
'   Сохранённые объекты: запросы qry_ВыручкаПоКатегориям, qry_ЗаказыКЭкспорту,
'   qry_ЗаказыМесяц (шаблон для блока 4) и отчёт rpt_Топ10Товаров.
'   Файлы создаются в папке базы (CurrentProject.Path).
'   В References должна стоять библиотека Microsoft Office XX.0 Access database engine
'   Object Library (DAO) — она включена по умолчанию.
'
' ПРИМЕЧАНИЕ ДЛЯ ACCESS 2007 RTM
'   Констант acFormatXLSX и acFormatPDF там может не быть (появились в SP2 / 2010).
'   Замените их строками "Excel Workbook (*.xlsx)" и "PDF Format (*.pdf)".
' =====================================================================================

Option Compare Database
Option Explicit

' ---------------------------------------------------------------------------------------------
' 1. Экспорт запроса в XLSX
' ---------------------------------------------------------------------------------------------
Public Sub ExportCategoryToXlsx()
    ' Экспортирует сохранённый запрос qry_ВыручкаПоКатегориям в файл Excel.
    ' Ожидаемый результат: файл qry_ВыручкаПоКатегориям.xlsx в папке базы
    ' с 5 строками категорий (bed_bath_table 769,50 ... auto 238,00 — см. aggregate-queries.sql).

    Dim strFile As String                 ' полный путь к файлу результата

    On Error GoTo ErrHandler              ' при любой ошибке исполнения переходим в обработчик

    strFile = CurrentProject.Path & "\qry_ВыручкаПоКатегориям.xlsx"   ' папка базы + имя файла
    DoCmd.OutputTo acOutputQuery, "qry_ВыручкаПоКатегориям", acFormatXLSX, strFile, False
    ' acOutputQuery — экспортируем запрос; acFormatXLSX — формат .xlsx;
    ' последний аргумент False — не открывать Excel автоматически после экспорта.

    MsgBox "Файл сохранён: " & strFile, vbInformation, "Экспорт в Excel"

    Exit Sub                              ' штатный выход — ОБЯЗАТЕЛЬНО выше обработчика

ErrHandler:                               ' обработчик ошибок
    ' Типичные причины: запрос не существует (2501/7874), файл занят Excel (открыт пользователем)
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Экспорт в Excel"
End Sub

' ---------------------------------------------------------------------------------------------
' 2. Экспорт отчёта в PDF
' ---------------------------------------------------------------------------------------------
Public Sub ExportReportToPdf()
    ' Экспортирует отчёт rpt_Топ10Товаров в PDF с сохранением вёрстки, группировок и итогов.
    ' Отличие от экспорта запроса: OutputTo для отчёта сохраняет форматирование и разрывы страниц.

    Dim strFile As String

    On Error GoTo ErrHandler

    strFile = CurrentProject.Path & "\rpt_Топ10Товаров.pdf"
    DoCmd.OutputTo acOutputReport, "rpt_Топ10Товаров", acFormatPDF, strFile, False
    ' acOutputReport — объект-отчёт; acFormatPDF — формат PDF; False — не открывать PDF самим.

    MsgBox "PDF сохранён: " & strFile, vbInformation, "Экспорт отчёта"

    Exit Sub

ErrHandler:
    ' Если отчёт не существует, Access вернёт ошибку 2501 («операция прервана») или 7874.
    ' В Access 2007 без надстройки Save as PDF экспорт в PDF недоступен — установите надстройку.
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Экспорт отчёта"
End Sub

' ---------------------------------------------------------------------------------------------
' 3. Экспорт запроса в CSV: со спецификацией и без неё
' ---------------------------------------------------------------------------------------------
Public Sub ExportQueryToCsv(Optional ByVal strSpec As String = "")
    ' Экспортирует qry_ЗаказыКЭкспорту в текстовый файл CSV.
    ' Параметр strSpec — имя СОХРАНЁННОЙ спецификации экспорта:
    '   ExportQueryToCsv          — без спецификации (настройки Access по умолчанию)
    '   ExportQueryToCsv "ЭкспортCSV" — с сохранённой спецификацией
    ' Спецификация создаётся один раз вручную: экспорт через мастер текстового файла,
    ' кнопка «Дополнительно» (разделитель, квалификатор, кодировка) → «Сохранить как...».
    ' Без спецификации Access пишет: разделитель запятая, первая строка — имена полей,
    ' кодировка ANSI (Windows-1251). Для «точка с запятой» и UTF-8 нужна своя спецификация.

    Dim strFile As String

    On Error GoTo ErrHandler

    strFile = CurrentProject.Path & "\orders_export.csv"

    If Len(strSpec) > 0 Then
        ' Вариант А: со спецификацией — второй аргумент задаёт её имя
        DoCmd.TransferText acExportDelim, strSpec, "qry_ЗаказыКЭкспорту", strFile, True
    Else
        ' Вариант Б: без спецификации — второй аргумент пустой
        DoCmd.TransferText acExportDelim, , "qry_ЗаказыКЭкспорту", strFile, True
    End If
    ' acExportDelim — текст с разделителями; True — первая строка содержит имена полей.

    MsgBox "CSV сохранён: " & strFile & " (строк: 12)", vbInformation, "Экспорт CSV"
    ' Ожидаемый результат: 12 строк данных + строка заголовков.

    Exit Sub

ErrHandler:
    ' Ошибка 3027/3051 — файл открыт или недоступна папка; 3625 — спецификация не найдена.
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Экспорт CSV"
End Sub

' ---------------------------------------------------------------------------------------------
' 4. Цикл по месяцам: отдельный CSV на каждый месяц заказов
' ---------------------------------------------------------------------------------------------
Public Sub ExportAllMonthly()
    ' Читает список месяцев из Orders, для каждого месяца переписывает SQL шаблонного запроса
    ' qry_ЗаказыМесяц и выгружает его в файл orders_ГГГГ-ММ.csv.
    ' ПОДГОТОВКА: создайте запрос qry_ЗаказыМесяц (любой SELECT из Orders) и сохраните его.
    ' Ожидаемый результат: 12 файлов — orders_2017-03.csv, orders_2017-04.csv, orders_2017-06.csv,
    ' orders_2017-08.csv, orders_2017-10.csv, orders_2017-11.csv (1 строка — отменённый O-10006),
    ' orders_2018-01.csv ... orders_2018-06.csv.

    Dim db As DAO.Database                ' текущая база
    Dim qdf As DAO.QueryDef               ' шаблонный запрос, у которого меняем SQL
    Dim rs As DAO.Recordset               ' список месяцев
    Dim strFile As String                 ' путь очередного файла
    Dim lngFiles As Long                  ' счётчик созданных файлов

    On Error GoTo ErrHandler

    Set db = CurrentDb                                     ' текущая база данных
    Set qdf = db.QueryDefs("qry_ЗаказыМесяц")               ' получаем сохранённый запрос-шаблон

    ' Уникальные месяцы заказов; FORMAT даёт текст вида 2017-03
    Set rs = db.OpenRecordset( _
        "SELECT DISTINCT FORMAT(OrderDate, 'yyyy-mm') AS Месяц FROM Orders ORDER BY Месяц", _
        dbOpenSnapshot)

    Do While Not rs.EOF                                    ' пока есть строки списка месяцев
        ' Переписываем SQL шаблонного запроса под текущий месяц;
        ' значения месяца берём из поля Месяц через rs!Месяц
        qdf.SQL = "SELECT o.* FROM Orders AS o WHERE FORMAT(o.OrderDate, 'yyyy-mm') = '" & rs!Месяц & "';"

        strFile = CurrentProject.Path & "\orders_" & rs!Месяц & ".csv"
        DoCmd.TransferText acExportDelim, , "qry_ЗаказыМесяц", strFile, True
        ' Выгружаем УЖЕ переписанный запрос в файл этого месяца

        lngFiles = lngFiles + 1
        rs.MoveNext                                        ' следующий месяц — без этого цикл зависнет
    Loop

    rs.Close                                               ' закрываем recordset
    Set rs = Nothing

    MsgBox "Создано файлов: " & lngFiles, vbInformation, "Экспорт по месяцам"
    ' Ожидаемое значение счётчика: 12.

    Exit Sub

ErrHandler:
    ' Частая причина: запрос qry_ЗаказыМесяц не создан (ошибка 3265 «элемент не найден»).
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Экспорт по месяцам"
    If Not rs Is Nothing Then rs.Close
End Sub
