Attribute VB_Name = "modCalendarGenerator"
' =====================================================================================
' modCalendarGenerator — генератор таблицы Calendar (файл к курсу
' «Microsoft Access 2007–2016 для аналитика», модуль 04).
'
' НАЗНАЧЕНИЕ
'   Три готовые процедуры:
'     CreateCalendarTable        — создаёт таблицу Calendar и заполняет всеми датами
'                                  периода 01.10.2016 — 31.10.2018 (761 запись)
'                                  через Recordset: рекомендуемый способ
'     CreateCalendarByInsertLoop — тот же результат циклом INSERT INTO: учебный
'                                  вариант, показывает синтаксис даты #...# в SQL,
'                                  работает заметно медленнее
'     CheckCalendar              — контрольные числа после заполнения (в Immediate)
'
' КАК ПОДКЛЮЧИТЬ
'   VBE (Alt+F11) → File → Import File... → выберите этот файл .bas;
'   либо Insert → Module (обычный модуль) и вставьте текст целиком.
'   Модуль появится в Project Explorer (Ctrl+R) под именем modCalendarGenerator.
'   Если в редакторе видны кракозябры — пересохраните текст в ANSI (Windows-1251):
'   подробно об этом — vba/README.md.
'
' КАК ЗАПУСТИТЬ
'   Поставьте курсор внутрь CreateCalendarTable и нажмите F5 (Run → Run Sub/UserForm).
'   Появится окно «Таблица Calendar создана: 761 записей…».
'   ВАЖНО: запускайте генератор ДО создания связи Calendar → Orders (в инструкции
'   импорта это шаг 6, а связь — шаг 8). Существующая связь или открытая таблица
'   Calendar блокируют её удаление и пересоздание.
'
' ТРЕБОВАНИЯ
'   База ShopData.accdb (или любая учебная база); библиотека Microsoft Office XX.0
'   Access database engine Object Library (DAO) включена по умолчанию.
'   Русские названия месяцев зашиты в код через Choose/Array, поэтому результат
'   НЕ зависит от языка Windows и региональных настроек.
' =====================================================================================

Option Compare Database
Option Explicit

' Границы календаря = период датасета Olist. Литералы дат в VBA всегда
' пишутся в американском формате #месяц/день/год#: #10/1/2016# = 1 октября 2016.
Private Const CAL_START As Date = #10/1/2016#
Private Const CAL_END As Date = #10/31/2018#

' ---------------------------------------------------------------------------------------------
' 1. Основной генератор: DDL + Recordset (рекомендуемый способ)
' ---------------------------------------------------------------------------------------------
Public Sub CreateCalendarTable()
    ' Создаёт таблицу Calendar заново и заполняет всеми датами периода.
    ' Ожидаемый результат: MsgBox «761 записей»; CheckCalendar подтверждает числа.

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim d As Date
    Dim lngWD As Long           ' день недели: 1 = понедельник … 7 = воскресенье
    Dim lngRows As Long         ' счётчик добавленных строк

    On Error GoTo ErrHandler

    Set db = CurrentDb

    ' Шаг 1. Удаляем старую версию таблицы, если она уже создавалась.
    On Error Resume Next
    db.TableDefs.Delete "Calendar"        ' ошибка игнорируется, если таблицы нет
    On Error GoTo ErrHandler

    ' Шаг 2. Создаём таблицу одним DDL-оператором.
    ' Типы: DATETIME = Дата/время; LONG = Числовой (целое); TEXT(n) = Короткий текст;
    ' YESNO = Логический; CONSTRAINT ... PRIMARY KEY = первичный ключ по CalDate.
    db.Execute "CREATE TABLE [Calendar] (" & _
        "[CalDate] DATETIME CONSTRAINT PK_Calendar PRIMARY KEY, " & _
        "[CalYear] LONG, [CalQuarter] LONG, [CalMonth] LONG, " & _
        "[MonthName] TEXT(20), [MonthYear] TEXT(15), " & _
        "[WeekDay] LONG, [IsWeekend] YESNO);", dbFailOnError

    ' Шаг 3. Заполняем даты через Recordset — пакетная вставка без диалогов подтверждения.
    Set rs = db.OpenRecordset("Calendar", dbOpenDynaset)
    For d = CAL_START To CAL_END            ' Date + 1 в цикле For = следующий день
        lngWD = Weekday(d, vbMonday)        ' vbMonday: понедельник = 1 … воскресенье = 7
        rs.AddNew
        rs!CalDate = d
        rs!CalYear = Year(d)
        rs!CalQuarter = (Month(d) + 2) \ 3  ' месяцы 1–3 → квартал 1 … 10–12 → квартал 4
        rs!CalMonth = Month(d)
        rs!MonthName = Choose(Month(d), "Январь", "Февраль", "Март", "Апрель", _
            "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
        rs!MonthYear = Choose(Month(d), "янв", "фев", "мар", "апр", "май", "июн", _
            "июл", "авг", "сен", "окт", "ноя", "дек") & " " & Year(d)
        rs!WeekDay = lngWD
        rs!IsWeekend = (lngWD >= 6)         ' суббота и воскресенье — выходной
        rs.Update
        lngRows = lngRows + 1
    Next d
    rs.Close

    MsgBox "Таблица Calendar создана: " & lngRows & " записей за период " & _
        Format(CAL_START, "dd.mm.yyyy") & " — " & Format(CAL_END, "dd.mm.yyyy") & ".", _
        vbInformation, "Генератор календаря"

    Exit Sub                              ' штатный выход — ОБЯЗАТЕЛЬНО выше обработчика

ErrHandler:                               ' обработчик ошибок
    ' Типичные причины: таблица Calendar открыта (ошибка 3211 «таблица занята»),
    ' существует связь Calendar → Orders — удалите её в Схеме данных и повторите.
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
        "Частые причины: открыта таблица Calendar (закройте её); " & _
        "создана связь Calendar → Orders (удалите связь в Схеме данных).", _
        vbCritical, "Генератор календаря"
End Sub

' ---------------------------------------------------------------------------------------------
' 2. Учебный вариант: тот же результат циклом INSERT INTO
' ---------------------------------------------------------------------------------------------
Public Sub CreateCalendarByInsertLoop()
    ' Каждая дата вставляется отдельным INSERT — так вы видите синтаксис литералов:
    ' текст в двойных кавычках, дата в решётках #...#, логическое «истина» как -1.
    ' На 761 строке работает в разы медленнее CreateCalendarTable — поэтому
    ' в реальной базе используйте процедуру 1.

    Dim db As DAO.Database
    Dim d As Date
    Dim lngWD As Long
    Dim lngRows As Long
    Dim strMonths As Variant          ' массив русских названий месяцев
    Dim strAbbr As Variant            ' массив трёхбуквенных сокращений

    strMonths = Array("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", _
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
    strAbbr = Array("янв", "фев", "мар", "апр", "май", "июн", _
        "июл", "авг", "сен", "окт", "ноя", "дек")

    On Error GoTo ErrHandler
    Set db = CurrentDb

    On Error Resume Next
    db.TableDefs.Delete "Calendar"
    On Error GoTo ErrHandler

    db.Execute "CREATE TABLE [Calendar] (" & _
        "[CalDate] DATETIME CONSTRAINT PK_Calendar PRIMARY KEY, " & _
        "[CalYear] LONG, [CalQuarter] LONG, [CalMonth] LONG, " & _
        "[MonthName] TEXT(20), [MonthYear] TEXT(15), " & _
        "[WeekDay] LONG, [IsWeekend] YESNO);", dbFailOnError

    For d = CAL_START To CAL_END
        lngWD = Weekday(d, vbMonday)
        ' Format(d, "yyyy\-mm\-dd") даёт дату вида 2016-10-01 независимо от локали;
        ' логическое поле принимает -1 (истина) и 0 (ложь) — это надёжнее,
        ' чем склеивать с SQL слово True, которое в русской локали станет «Истина».
        db.Execute "INSERT INTO [Calendar] " & _
            "([CalDate],[CalYear],[CalQuarter],[CalMonth],[MonthName],[MonthYear],[WeekDay],[IsWeekend]) " & _
            "VALUES (#" & Format(d, "yyyy\-mm\-dd") & "#, " & Year(d) & ", " & _
            ((Month(d) + 2) \ 3) & ", " & Month(d) & ", " & _
            """" & strMonths(Month(d) - 1) & """, " & _
            """" & strAbbr(Month(d) - 1) & " " & Year(d) & """, " & _
            lngWD & ", " & IIf(lngWD >= 6, -1, 0) & ");", dbFailOnError
        lngRows = lngRows + 1
    Next d

    MsgBox "Таблица Calendar создана циклом INSERT: " & lngRows & " записей.", _
        vbInformation, "Генератор календаря"

    Exit Sub

ErrHandler:
    MsgBox "Ошибка " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
        "Частые причины: открыта таблица Calendar (закройте её); " & _
        "создана связь Calendar → Orders (удалите связь в Схеме данных).", _
        vbCritical, "Генератор календаря"
End Sub

' ---------------------------------------------------------------------------------------------
' 3. Проверка после заполнения: числа появятся в окне Immediate (Ctrl+G)
' ---------------------------------------------------------------------------------------------
Public Sub CheckCalendar()
    ' Запустите курсором внутри процедуры (F5) и откройте Immediate (Ctrl+G).
    ' Ожидаемые значения показаны в комментариях справа.

    Debug.Print "Всего дат: " & DCount("*", "Calendar")                        ' 761
    Debug.Print "Первая дата: " & Format(DMin("CalDate", "Calendar"), "dd.mm.yyyy")   ' 01.10.2016
    Debug.Print "Последняя дата: " & Format(DMax("CalDate", "Calendar"), "dd.mm.yyyy") ' 31.10.2018
    Debug.Print "Дат 2016 года: " & DCount("*", "Calendar", "[CalYear] = 2016")  ' 92
    Debug.Print "Дат 2017 года: " & DCount("*", "Calendar", "[CalYear] = 2017")  ' 365
    Debug.Print "Дат 2018 года: " & DCount("*", "Calendar", "[CalYear] = 2018")  ' 304
    Debug.Print "Выходных: " & DCount("*", "Calendar", "[IsWeekend] = True")     ' 218
    Debug.Print "Будних: " & DCount("*", "Calendar", "[IsWeekend] = False")      ' 543
    Debug.Print "01.10.2016 — это: " & DLookup("MonthName", "Calendar", _
        "[CalDate] = #10/1/2016#") & ", день недели " & _
        DLookup("WeekDay", "Calendar", "[CalDate] = #10/1/2016#")                ' Октябрь, 6 (сб)
End Sub
