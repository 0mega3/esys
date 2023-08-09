VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "HCI2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Dim byteTrigger As Byte

'The function returns the address of the cell.
Private Function GetCellAddress(strWorksheetName As String, _
                                intRow As Integer, _
                                intColumn As Integer, _
                                strParamName As String)
    Dim i As Integer, j As Integer
    For i = 1 To 100
        For j = 1 To 100
            If IsError(Worksheets(strWorksheetName).Cells(i, j).Value) = False Then
                If Worksheets(strWorksheetName).Cells(i, j).Value = strParamName Then
                    intRow = i
                    intColumn = j
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

'The function returns an array of expiration dates of options.
Private Function varGetExpDateArray(ByVal strWorksheetName As String, _
                                    ByVal intCapRow As Integer, _
                                    ByVal intCapColumn As Integer, _
                                    ByVal strLifespan As String) As Variant
    Dim i As Integer, j As Integer
    Dim intLastDateRow As Integer
    Dim varExpDateArray() As Variant

    Dim rngDateRange As Range, rngCell As Range, _
        varValueCell As Variant
    Set rngDateRange = Worksheets(strWorksheetName).Range("A1:A20000")

    For Each rngCell In rngDateRange
        varValueCell = rngCell.Value
        If IsEmpty(varValueCell) = True Then
            intLastDateRow = rngCell.Row - 1
            Exit For
        End If
    Next
    If UCase(strLifespan) = "Q" Then
        ReDim varExpDateArray((intLastDateRow - intCapRow) / 3)
        j = 1
        For i = 3 To intLastDateRow Step 3
            varValueCell = Worksheets(strWorksheetName).Cells(intCapRow + i, intCapColumn).Value
            If IsDate(varValueCell) Then
                varExpDateArray(j) = varValueCell
                j = j + 1
            End If
        Next i
    Else
        Dim colExpDates As Collection
        Set colExpDates = New Collection
        For i = 1 To (intLastDateRow - intCapRow)
            varValueCell = Worksheets(strWorksheetName).Cells(intCapRow + i, intCapColumn).Value
            If IsEmpty(varValueCell) = False Then
                If UCase(strLifespan) = "M" Then
                    Dim strCap As String, strFutType As String
                    strCap = Worksheets(strWorksheetName).Cells(intCapRow, intCapColumn).Value
                    strFutType = Left(strCap, 3)
                    If UCase(strFutType) = "FXF" Then
                        If IsDate(varValueCell) Then
                            If CDate(varValueCell) >= CDate("16.09.2013") Then
                                colExpDates.Add (varValueCell)
                            End If
                        End If
                    Else
                        colExpDates.Add (varValueCell)
                    End If
                ElseIf UCase(strLifespan) = "W" Then
                    'The split function returns a zero-based, one-dimensional array
                    varValueCell = Split(varValueCell, ChrW(44)) 'Comma delimiter
                    For j = 0 To UBound(varValueCell)
                        If IsDate(varValueCell(j)) Then
                            colExpDates.Add (varValueCell(j))
                        End If
                    Next j
                End If
            End If
        Next i
        ReDim varExpDateArray(colExpDates.Count)
        For i = 1 To UBound(varExpDateArray)
            varExpDateArray(i) = colExpDates(i)
        Next i
        Set colExpDates = Nothing
    End If
    varGetExpDateArray = varExpDateArray
End Function

'Функция дает количество файл-отчетов в папке назначения.
Private Function intGetQtyReportFiles(varDirectory As Variant) As Integer
    Dim strReportFileName As String
    Dim intCounter As Integer
    'Получение имени первого файл-отчета (образец первого вызова).
    strReportFileName = Dir(varDirectory, vbReadOnly + vbHidden + vbSystem)
    If strReportFileName = "" Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "В папке назначения файл-отчетов нет." _
        & vbNewLine & "Поместите файл-отчеты в папку и повторите расчет IV.")
        Exit Function
    Else
        'Запуск счетчика просмотров.
        intCounter = 1
    End If
    'Получение имен остальных файл-отчетов (по образцу первого вызова).
    Do While strReportFileName <> ""
        strReportFileName = Dir()
        intCounter = intCounter + 1
    Loop
    intGetQtyReportFiles = intCounter - 1 'Вычитаем последний просмотр (содержит "ПУСТО").
End Function

'Функция дает массив имен файл-отчетов.
Private Function varGetReportFileNameArray(varDirectory As Variant, _
                                           intQtyReportFiles As Integer) As Variant
    Dim strReportFileNameArray() As String
    Dim intCounter As Integer
    ReDim strReportFileNameArray(intQtyReportFiles)
    'Получение имени первого файл-отчета (образец первого вызова).
    strReportFileNameArray(1) = Dir(varDirectory, vbReadOnly + vbHidden + vbSystem)
    'Запуск счетчика записи элементов.
    intCounter = 1
    'Получение имен остальных файл-отчетов (по образцу первого вызова).
    Do While intQtyReportFiles <> intCounter
        intCounter = intCounter + 1
        strReportFileNameArray(intCounter) = Dir()
    Loop
    varGetReportFileNameArray = strReportFileNameArray
End Function

'Функция ловит дату, используя регулярное выражение.
'В VBA Excel для работы с регулярными выражениями используется библиотека
'«Microsoft VBScript Regular Expression».
'Создание объекта RegExp:
'Раннее связывание заключается в присвоении нового экземпляра объекта RegExp
'переменной, уже объявленной, как переменная определенного типа (в нашем случае,
'как RegExp).
'Для осуществления ранней привязки необходимо подключить к проекту VBA ссылку
'на библиотеку «Microsoft VBScript Regular Expression», для чего в редакторе VBA
'выбираем Tools - References...
'В открывшемся окне «References» находим строку «Microsoft VBScript Regular
'Expression 5.5» (если у вас ее нет, то строку «Microsoft VBScript Regular
'Expression 1.0»), отмечаем ее галочкой и нажимаем «ОК».
'Позднее связывание заключается в присвоении нового экземпляра объекта RegExp
'переменной, объявленной как Object, с помощью функции CreateObject.
Private Function varGetDateFromReportFileName(varReportFileName As Variant) As Variant
    Dim strReDate As String
    'Dim Re As Object
    'Set Re = CreateObject("vbscript.regexp")
    Dim Re As RegExp
    Set Re = New RegExp
    'Re.Pattern = "\d+( |.)([яфмаисонд](([а-я]{2}\.? )|([а-я]+[а|я] ))\d{4})|(\d+( |.)\d+)"
    'It will match on well-formed but invalid values like 20010231 (31th Feb!) :)
    Re.Pattern = "(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12]\d|30|31)" 'YYYYMMDD string with no delimiters
    If Re.test(varReportFileName) Then strReDate = Re.Execute(varReportFileName)(0)
    Set Re = Nothing
    varGetDateFromReportFileName = DateSerial(Left(strReDate, 4), Mid(strReDate, 5, 2), Right(strReDate, 2))
End Function

'Функция создает массив дат файл-отчетов.
Private Function varGetDateArrayFromReportFileNames(varReportFileNameArray As Variant) As Variant
    Dim i As Integer
    Dim varDateArrayFromReportFileNames() As Variant
    ReDim varDateArrayFromReportFileNames(UBound(varReportFileNameArray))
    For i = 1 To UBound(varReportFileNameArray)
        varDateArrayFromReportFileNames(i) = varGetDateFromReportFileName(varReportFileNameArray(i))
    Next i
    varGetDateArrayFromReportFileNames = varDateArrayFromReportFileNames
End Function

'Функция считает количество дат, выловленных из массива имен файл-отчетов.
Private Function intGetQtyDatesFromReportFileNames(varDateArrayFromReportFileNames As Variant) As Integer
    Dim i As Integer, intCounter As Integer
    intCounter = 0
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        If IsDate(varDateArrayFromReportFileNames(i)) = True Then
            intCounter = intCounter + 1
        End If
    Next i
    intGetQtyDatesFromReportFileNames = intCounter
End Function

'Функция сортирует массив дат в связке с массивом имен файл-отчетов по возрастанию
'(метод Пузырек).
Private Function BubbleSort(varReportFileNameArray As Variant, _
                            varDateArrayFromReportFileNames As Variant)
    Dim i As Integer
    Dim boolSwitch As Boolean
    Dim varTmp As Variant
    Do
        boolSwitch = 0
        For i = 1 To UBound(varDateArrayFromReportFileNames) - 1
            If DateValue(varDateArrayFromReportFileNames(i + 1)) < DateValue(varDateArrayFromReportFileNames(i)) Then
                boolSwitch = 1
                varTmp = varDateArrayFromReportFileNames(i)
                varDateArrayFromReportFileNames(i) = varDateArrayFromReportFileNames(i + 1)
                varDateArrayFromReportFileNames(i + 1) = varTmp

                varTmp = varReportFileNameArray(i)
                varReportFileNameArray(i) = varReportFileNameArray(i + 1)
                varReportFileNameArray(i + 1) = varTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
End Function

'Функция сортирует массив страйков в связке с массивом волатильностей и объемов
'по возрастанию (метод Пузырек).
Private Function BubbleSort2(varStrikeArray As Variant, _
                             varIVArray As Variant, _
                             varVolArray As Variant)
    Dim i As Integer
    Dim boolSwitch As Boolean
    Dim varTmp As Variant

    Do
        boolSwitch = 0
        For i = 1 To UBound(varStrikeArray) - 1
            If varStrikeArray(i + 1) < varStrikeArray(i) Then
                boolSwitch = 1
                varTmp = varStrikeArray(i)
                varStrikeArray(i) = varStrikeArray(i + 1)
                varStrikeArray(i + 1) = varTmp

                varTmp = varIVArray(i)
                varIVArray(i) = varIVArray(i + 1)
                varIVArray(i + 1) = varTmp

                varTmp = varVolArray(i)
                varVolArray(i) = varVolArray(i + 1)
                varVolArray(i + 1) = varTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
End Function

'Функция приводит массив в соответствие с оператором Option Base 1.
Private Function varToOptionBase1(varArrayOptionBase0 As Variant) As Variant
    Dim i As Integer, j As Integer, _
    varArrayOptionBase1() As Variant
    ReDim varArrayOptionBase1(UBound(varArrayOptionBase0) + 1)
    j = 1
    For i = 0 To UBound(varArrayOptionBase0)
        varArrayOptionBase1(j) = varArrayOptionBase0(i)
        j = j + 1
    Next i
    varToOptionBase1 = varArrayOptionBase1
End Function

'Функция проверяет, чтобы пара файл-отчетов pr-options-delta-ru и rts-results-ru
'имела одинаковую дату.
Private Function boolGetCheckDualDates(varDateArrayFromReportFileNames As Variant) As Boolean
    Dim i As Integer
    For i = 1 To UBound(varDateArrayFromReportFileNames) Step 2 'Сравниваем 1 и 2, 3 и 4, 5 и 6, и т.д.
        If varDateArrayFromReportFileNames(i) <> varDateArrayFromReportFileNames(i + 1) Then
            boolGetCheckDualDates = False
            Exit Function
        End If
    Next i
    boolGetCheckDualDates = True
End Function

'Функция фильтрует повторяющиеся даты.
Private Function varGetFilterDateArrayFromReportFileNames(varDateArrayFromReportFileNames As Variant) As Variant
    Dim i As Integer, intCounter As Integer
    Dim varFilterDateArrayFromReportFileNames() As Variant
    ReDim varFilterDateArrayFromReportFileNames(UBound(varDateArrayFromReportFileNames) / 2)
    intCounter = 1
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        If i = 1 Then 'Если дата первая.
            varFilterDateArrayFromReportFileNames(intCounter) = varDateArrayFromReportFileNames(i)
        ElseIf varDateArrayFromReportFileNames(i) <> varDateArrayFromReportFileNames(i - 1) Then 'Если текущая дата не равна предыдущей.
            intCounter = intCounter + 1
            varFilterDateArrayFromReportFileNames(intCounter) = varDateArrayFromReportFileNames(i)
        End If
    Next i
    varGetFilterDateArrayFromReportFileNames = varFilterDateArrayFromReportFileNames
End Function

'Функция дает массив цен фьючерса.
Private Function varGetFuturesPriceArray(varProductCode As Variant, _
                                         varDateArrayFromReportFileNames As Variant) As Variant
    Dim intDateCounter As Integer, lngRowCounter As Long
    Dim varWorkSheet As Variant
    Dim rngDateRange_HV As Range, rngCell As Range, varValueCell As Variant
    Dim varFuturesPriceArray() As Variant
    ReDim varFuturesPriceArray(UBound(varDateArrayFromReportFileNames))
    'Назначаем рабочий лист и диапазон дат HV.
    If varProductCode = "SBRF" Then
        Set varWorkSheet = Worksheets("SBRF_HV")
        Set rngDateRange_HV = varWorkSheet.Range("A1:A20000")
    ElseIf varProductCode = "Si" Then
        Set varWorkSheet = Worksheets("Si_HV")
        Set rngDateRange_HV = varWorkSheet.Range("A1:A20000")
    ElseIf varProductCode = "RTS" Then
        Set varWorkSheet = Worksheets("RI_HV")
        Set rngDateRange_HV = varWorkSheet.Range("A1:A20000")
    ElseIf varProductCode = "BR" Then
        Set varWorkSheet = Worksheets("BR_HV")
        Set rngDateRange_HV = varWorkSheet.Range("A1:A20000")
    ElseIf varProductCode = "GOLD" Then
        Set varWorkSheet = Worksheets("GD_HV")
        Set rngDateRange_HV = varWorkSheet.Range("A1:A20000")
    End If
    intDateCounter = 1 'Счетчик дат.
    lngRowCounter = 0 'Счетчик строк.
    For Each rngCell In rngDateRange_HV 'Начинаем просмотр каждой ячейки диапазона.
        varValueCell = rngCell.Value
        lngRowCounter = lngRowCounter + 1
        If varValueCell = CDate(varDateArrayFromReportFileNames(intDateCounter)) Then 'Если текущее значение ячейки (дата) совпадает с датой из массива дат.
            varFuturesPriceArray(intDateCounter) = varWorkSheet.Cells(lngRowCounter, 2).Value 'Считываем значение цены с соседней ячейки справа.
            If intDateCounter = UBound(varFuturesPriceArray) Then 'Если массив цен заполнен полностью.
                varGetFuturesPriceArray = varFuturesPriceArray 'Возвращаем массив цен.
                Exit Function
            Else: intDateCounter = intDateCounter + 1 'Переходим к следующей дате из массива дат.
            End If
        End If
    Next
End Function

'Функция дает один массив дат экспирации относительно дат расчета (Front).
Private Function ClosestExpDatesNoCombo(varExpDateArray As Variant, _
                                        varDateArrayFromReportFileNames As Variant, _
                                        varFrontExpDateArray As Variant)
    Dim i As Integer, j As Integer
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        For j = 1 To UBound(varExpDateArray)
            If DateValue(CDate(varDateArrayFromReportFileNames(i))) < DateValue(varExpDateArray(j)) Then
                varFrontExpDateArray(i) = CStr(varExpDateArray(j))
                Exit For
            End If
            If DateValue(CDate(varDateArrayFromReportFileNames(i))) = DateValue(varExpDateArray(j)) Then
                varFrontExpDateArray(i) = CStr(varExpDateArray(j + 1))
                Exit For
            End If
        Next j
    Next i
End Function

'Функция дает два массива дат экспирации относительно дат расчета (Front->Next).
Private Function ClosestExpDatesCombo2(varExpDateArray As Variant, _
                                       varDateArrayFromReportFileNames As Variant, _
                                       varFrontExpDateArray As Variant, _
                                       varNextExpDateArray As Variant)
    Dim i As Integer, j As Integer
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        For j = 1 To UBound(varExpDateArray)
            If DateValue(CDate(varDateArrayFromReportFileNames(i))) < DateValue(varExpDateArray(j)) Then
                varFrontExpDateArray(i) = CStr(varExpDateArray(j))
                varNextExpDateArray(i) = CStr(varExpDateArray(j + 1))
                Exit For
            End If
            If DateValue(CDate(varDateArrayFromReportFileNames(i))) = DateValue(varExpDateArray(j)) Then
                varFrontExpDateArray(i) = CStr(varExpDateArray(j + 1))
                varNextExpDateArray(i) = CStr(varExpDateArray(j + 2))
                Exit For
            End If
        Next j
    Next i
End Function

'Функция дает три массива дат экспирации относительно дат расчета (Front->Next->Back).
Private Function ClosestExpDatesCombo3(varExpDateArray As Variant, _
                                       varDateArrayFromReportFileNames As Variant, _
                                       varFrontExpDateArray As Variant, _
                                       varNextExpDateArray As Variant, _
                                       varBackExpDateArray As Variant)
    Dim i As Integer, j As Integer
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        For j = 1 To UBound(varExpDateArray)
            If DateValue(CDate(varDateArrayFromReportFileNames(i))) < DateValue(varExpDateArray(j)) Then
                varFrontExpDateArray(i) = CStr(varExpDateArray(j))
                varNextExpDateArray(i) = CStr(varExpDateArray(j + 1))
                varBackExpDateArray(i) = CStr(varExpDateArray(j + 2))
                Exit For
            End If
            If DateValue(CDate(varDateArrayFromReportFileNames(i))) = DateValue(varExpDateArray(j)) Then
                varFrontExpDateArray(i) = CStr(varExpDateArray(j + 1))
                varNextExpDateArray(i) = CStr(varExpDateArray(j + 2))
                varBackExpDateArray(i) = CStr(varExpDateArray(j + 3))
                Exit For
            End If
        Next j
    Next i
End Function

'Функция считывает содержимое файл-отчета в строку.
Private Function strGetContentFromReportFile(varDirectory As Variant, _
                                             varReportFileName As Variant) As String
    Dim intReportFileNum As Integer, _
        strContentFromReportFile As String
    intReportFileNum = FreeFile
    Open varDirectory & varReportFileName For Input As intReportFileNum
    While Not EOF(intReportFileNum)
        Line Input #intReportFileNum, strContentFromReportFile
    Wend
    Close #intReportFileNum
    strGetContentFromReportFile = strContentFromReportFile
End Function

'Функция проверяет наличие даты экспирации в содержимом файл-отчета. UNUSED
Private Function boolIsThereExpDate(varContentFromReportFile As Variant, _
                                    varFrontExpDateArray As Variant) As Boolean
    Dim strContentStringFromReportFile As String, _
        Re As RegExp
    Set Re = New RegExp
    With Re
        .Global = False 'До первого совпадения.
        .IgnoreCase = False 'Учитывать регистр.
        .MultiLine = False 'Однострочный объект.
        .Pattern = Format(varFrontExpDateArray, "yyyy-mm-dd") 'Шаблон.
    End With
    strContentStringFromReportFile = Join(varContentFromReportFile, vbLf) 'Собираем содержимое в строку.
    If Re.test(strContentStringFromReportFile) Then 'Если есть совпадение с шаблоном.
        boolIsThereExpDate = True
    Else: boolIsThereExpDate = False 'Если совпадения нет.
    End If
    Set Re = Nothing
End Function

'Returns an array which exactly matches this collection.
' Note: This function is not safe for concurrent modification.
Private Function varToArray(col As Collection) As Variant
    Dim i As Long, _
        varA() As Variant
    ReDim varA(0 To col.Count)
    For i = 0 To col.Count - 1
        varA(i) = col(i + 1)
    Next i
    varToArray = varA()
End Function

'Advanced Filter
Private Function varAdvFilter(ByVal varArray As Variant, _
                              ByVal strMatch As String) As Variant
    Dim i As Long, _
        col As Collection
    Set col = New Collection
    'varArray is a zero-based array.
    'Traversing the array starts with the 1st element,
    'since the cap is ignored.
    For i = 1 To UBound(varArray)
        If strMatch = Mid(varArray(i), 1, Len(strMatch)) Then col.Add (varArray(i))
    Next i
    varAdvFilter = varToArray(col)
End Function

'Функция дает Call и Put массивы страйков и волатильностей из файл-отчета pr-options-delta-ru.
Private Function StrikeAndIVArrays(varDirectory As Variant, _
                                   varReportFileNameArray_pr_options_delta_ru As Variant, _
                                   varProductCode As Variant, _
                                   varFrontExpDateArray As Variant, _
                                   varCallStrikeArray() As Variant, _
                                   varPutStrikeArray() As Variant, _
                                   varCallIVArray() As Variant, _
                                   varPutIVArray() As Variant, _
                                   boolIsThereExpDate_pr_options_delta_ru As Boolean)
    Dim i As Integer, _
        varContentFromReportFile As Variant, _
        varCallContentFromReportFile As Variant, _
        varPutContentFromReportFile As Variant, _
        varContentFilterFromReportFile As Variant
    'Читаем файл-отчет pr-options-delta-ru.
    varContentFromReportFile = strGetContentFromReportFile(varDirectory, varReportFileNameArray_pr_options_delta_ru) 'Считываем содержимое файл-отчета в строку.
    varContentFromReportFile = Split(varContentFromReportFile, vbLf) 'Разбиваем содержимое файл-отчета на строки. Разделитель - символ LF - Chr(10) - vbLf.
    varContentFromReportFile = Filter(varContentFromReportFile, ";" & varProductCode & "-") 'Фильтруем содержимое относительно торгового инструмента.
    'boolIsThereExpDate_pr_options_delta_ru = boolIsThereExpDate(varContentFromReportFile, varFrontExpDateArray) 'Проверяем наличие даты экспирации. UNUSED
    varContentFromReportFile = Filter(varContentFromReportFile, Format(varFrontExpDateArray, "yyyy-mm-dd")) 'Фильтруем содержимое относительно даты экспирации.
    boolIsThereExpDate_pr_options_delta_ru = True 'False by default
    If UBound(varContentFromReportFile) = -1 Then 'Если даты экспирации нет.
        boolIsThereExpDate_pr_options_delta_ru = False
        Exit Function
    Else
        'Волатильность Call и Put опционов на одном страйке одинаковая, но кол-во Call страйков может быть <> кол-ву Put страйков.
        varCallContentFromReportFile = Filter(varContentFromReportFile, "Call") 'Фильтруем содержимое относительно опционов Call.
        varPutContentFromReportFile = Filter(varContentFromReportFile, "Put") 'Фильтруем содержимое относительно опционов Put.
        varContentFromReportFile = Empty
        varCallContentFromReportFile = varToOptionBase1(varCallContentFromReportFile) 'Приводим массив в соответствие с оператором Option Base 1.
        varPutContentFromReportFile = varToOptionBase1(varPutContentFromReportFile)
        ReDim varCallStrikeArray(UBound(varCallContentFromReportFile))
        ReDim varPutStrikeArray(UBound(varPutContentFromReportFile))
        ReDim varCallIVArray(UBound(varCallContentFromReportFile))
        ReDim varPutIVArray(UBound(varPutContentFromReportFile))
        'Получаем Call массивы страйков и волатильностей
        For i = 1 To UBound(varCallContentFromReportFile)
            varContentFilterFromReportFile = Split(varCallContentFromReportFile(i), ";")
            varCallStrikeArray(i) = CLng(varContentFilterFromReportFile(4)) 'Т.к. функция Split определяет нижнюю границу массива=0, то 5 полю соот. 4.
            varCallIVArray(i) = CDbl(varContentFilterFromReportFile(9)) 'Т.к. функция Split определяет нижнюю границу массива=0, то 10 полю соот. 9.
        Next i
        'Получаем Put массивы страйков и волатильностей
        For i = 1 To UBound(varPutContentFromReportFile)
            varContentFilterFromReportFile = Split(varPutContentFromReportFile(i), ";")
            varPutStrikeArray(i) = CLng(varContentFilterFromReportFile(4))
            varPutIVArray(i) = CDbl(varContentFilterFromReportFile(9))
        Next i
    End If
End Function

'Функция дает Call и Put массивы объемов из файл-отчета rts-results-ru.
Private Function VolArrays(varDirectory As Variant, _
                           varReportFileNameArray_rts_results_ru As Variant, _
                           varProductCode As Variant, _
                           varFrontExpDateArray As Variant, _
                           varCallVolArray() As Variant, _
                           varPutVolArray() As Variant, _
                           boolIsThereExpDate_rts_results_ru As Boolean)
    Dim i As Integer, _
        varContentFromReportFile As Variant, _
        varCallContentFromReportFile As Variant, _
        varPutContentFromReportFile As Variant, _
        varContentFilterFromReportFile As Variant
    'Читаем файл-отчет rts-results-ru.
    varContentFromReportFile = strGetContentFromReportFile(varDirectory, varReportFileNameArray_rts_results_ru)
    varContentFromReportFile = Split(varContentFromReportFile, vbLf)
    varContentFromReportFile = varAdvFilter(varContentFromReportFile, varProductCode)
    'boolIsThereExpDate_rts_results_ru = boolIsThereExpDate(varContentFromReportFile, varFrontExpDateArray) 'Проверяем наличие даты экспирации. UNUSED
    varContentFromReportFile = Filter(varContentFromReportFile, varFrontExpDateArray)
    boolIsThereExpDate_rts_results_ru = True
    If UBound(varContentFromReportFile) = -1 Then 'Если даты экспирации нет.
        boolIsThereExpDate_rts_results_ru = False
        Exit Function
    Else
        varCallContentFromReportFile = Filter(varContentFromReportFile, "CA") 'Фильтруем содержимое относительно опционов Call.
        varPutContentFromReportFile = Filter(varContentFromReportFile, "PA") 'Фильтруем содержимое относительно опционов Put.
        varContentFromReportFile = Empty
        varCallContentFromReportFile = varToOptionBase1(varCallContentFromReportFile)
        varPutContentFromReportFile = varToOptionBase1(varPutContentFromReportFile)
        ReDim varCallVolArray(UBound(varCallContentFromReportFile))
        ReDim varPutVolArray(UBound(varPutContentFromReportFile))
        'Кол-во Call строк может быть <> кол-ву Put строк
        'Получаем массив Call объемов
        For i = 1 To UBound(varCallContentFromReportFile)
            varContentFilterFromReportFile = Split(varCallContentFromReportFile(i), ";")
            If varContentFilterFromReportFile(12) = "" Then 'Т.к. функция Split определяет нижнюю границу массива=0, то 13 полю соот. 12.
                varCallVolArray(i) = CLng(0) 'Если ничего, то значение объема=0.
            Else: varCallVolArray(i) = CLng(varContentFilterFromReportFile(12))
            End If
        Next i
        'Получем массив Put объемов
        For i = 1 To UBound(varPutContentFromReportFile)
            varContentFilterFromReportFile = Split(varPutContentFromReportFile(i), ";")
            If varContentFilterFromReportFile(12) = "" Then
                varPutVolArray(i) = CLng(0)
            Else: varPutVolArray(i) = CLng(varContentFilterFromReportFile(12))
            End If
        Next i
    End If
End Function
'-------------------------------------------------------------------------------
'Блок IV-начало.
'-------------------------------------------------------------------------------
'Функция выполняет поиск двух центральных страйков, проверкой условия
'на попадание цены в диапазон между соседними страйками.
'После того, как два центральных страйка найдены, функция сохраняет их номера
'и выполняет сравнение объемов на этих страйках.
'После сравнения объемов, Функция сохраняет значения рыночных волатильностей
'в зависимости от объема на этих страйках.
'Переменной varIVCSWithBigVol присваивается рыночная волатильность от страйка
'с большим объемом, а переменной varIVCSWithSmallVol - от страйка с меньшим объемом.
'Если на обоих центральных страйках объемы равны нулю, то переменной
'varIVCSWithBigVol присваивается рыночная волатильность от страйка с большим
'индексом (учитывается логнормальное распределение).
Private Function CentralStrike(varFuturesPrice As Variant, _
                               varVolArray As Variant, _
                               varStrikeArray As Variant, _
                               varIVArray As Variant, _
                               intNumCS1 As Integer, _
                               intNumCS2 As Integer, _
                               varIVCSWithBigVol As Variant, _
                               varIVCSWithSmallVol As Variant, _
                               varAvgValueCS1CS2 As Variant)
    Dim i As Integer
    For i = 1 To UBound(varStrikeArray) - 1
        If varFuturesPrice >= varStrikeArray(i) And varFuturesPrice <= varStrikeArray(i + 1) Then
            intNumCS1 = i
            intNumCS2 = i + 1
            varAvgValueCS1CS2 = (varStrikeArray(intNumCS1) + varStrikeArray(intNumCS2)) / 2 'Вычисляет среднее значение 2-х центральных страйков.
            If varVolArray(i) > varVolArray(i + 1) Then
                varIVCSWithBigVol = varIVArray(i)
                varIVCSWithSmallVol = varIVArray(i + 1)
                Exit For
            Else
                varIVCSWithBigVol = varIVArray(i + 1)
                varIVCSWithSmallVol = varIVArray(i)
                Exit For
            End If
        End If
    Next i
End Function

'Функция считает все ненулевые объемы, не затрагивая значения 2-х центральных
'страйков.
Private Function lngGetQtyNonZeroVol(varVolArray As Variant, _
                                     intNumCS1 As Integer, _
                                     intNumCS2 As Integer, _
                                     lngQtyNonZeroVol As Long) As Long
    Dim i As Integer, _
        intCount As Integer
    For i = 1 To UBound(varVolArray)
        If IsNumeric(varVolArray(i)) = True And varVolArray(i) > 0 And i <> intNumCS1 And i <> intNumCS2 Then
            lngQtyNonZeroVol = lngQtyNonZeroVol + 1
        End If
    Next i
    lngGetQtyNonZeroVol = lngQtyNonZeroVol
End Function

'Функция выбирает все ненулевые объемы, соответствующие им волатильности
'и страйки, не затрагивая значения 2-х центральных страйков, и сохраняет
'в массивах.
Private Function Selection(varVolArray As Variant, _
                           varStrikeArray As Variant, _
                           varIVArray As Variant, _
                           intNumCS1 As Integer, _
                           intNumCS2 As Integer, _
                           varSelectiveVolArray As Variant, _
                           varSelectiveStrikeArray As Variant, _
                           varSelectiveIVArray As Variant)
    Dim i As Integer, _
        intCount As Integer
    intCount = 0
    For i = 1 To UBound(varVolArray)
        If IsNumeric(varVolArray(i)) = True And varVolArray(i) > 0 And i <> intNumCS1 And i <> intNumCS2 Then
            intCount = intCount + 1
            varSelectiveVolArray(intCount) = varVolArray(i)
            varSelectiveStrikeArray(intCount) = varStrikeArray(i)
            varSelectiveIVArray(intCount) = varIVArray(i)
        End If
    Next i
End Function

'Функция сортирует массив рыночных волатильностей в зависимости от объема
'на страйках в связке с массивом объемов и массивом страйков по убыванию
'(метод Пузырек).
Private Function SortIV(varSelectiveVolArray As Variant, _
                        varSelectiveStrikeArray As Variant, _
                        varSelectiveIVArray As Variant)
    Dim boolSwitch As Boolean
    Dim i As Integer
    Dim dblTmp As Variant
    Do
        boolSwitch = 0
        For i = 1 To UBound(varSelectiveVolArray) - 1
            If varSelectiveVolArray(i + 1) > varSelectiveVolArray(i) Then
                boolSwitch = 1
                'Сортировка массива объемов.
                dblTmp = varSelectiveVolArray(i)
                varSelectiveVolArray(i) = varSelectiveVolArray(i + 1)
                varSelectiveVolArray(i + 1) = dblTmp
                'Сортировка массива страйков.
                dblTmp = varSelectiveStrikeArray(i)
                varSelectiveStrikeArray(i) = varSelectiveStrikeArray(i + 1)
                varSelectiveStrikeArray(i + 1) = dblTmp
                'Сортировка массива рыночных волатильностей.
                dblTmp = varSelectiveIVArray(i)
                varSelectiveIVArray(i) = varSelectiveIVArray(i + 1)
                varSelectiveIVArray(i + 1) = dblTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
End Function

'Функция выполняет точечную корректирующую сортировку массива рыночных
'волатильностей.
'Функция точечно сортирует массив расстояний от центра в связке с массивом
'рыночных волатильностей.
'Сортировка происходит пузырьковым методом от минимального значения
'к максимальному.
Private Function CorrectSortIV(varArrayLength As Variant, _
                               varSelectiveIVArray As Variant, _
                               intNumGroup As Integer, _
                               intBegin As Integer)
    Dim boolSwitch As Boolean
    Dim i As Integer, _
        j As Integer
    Dim dblTmp As Variant
    Do
        boolSwitch = 0
        If intNumGroup = 1 Then '/Условие определяющее точку отсчета.
            j = 1
        Else: j = intBegin
        End If
        For i = j To UBound(varArrayLength) - 1
            If varArrayLength(i + 1) < varArrayLength(i) Then
                boolSwitch = 1
                'Сортировка массива расстояний.
                dblTmp = varArrayLength(i)
                varArrayLength(i) = varArrayLength(i + 1)
                varArrayLength(i + 1) = dblTmp
                'Корректирующая сортировка массива рыночных волатильностей.
                dblTmp = varSelectiveIVArray(i)
                varSelectiveIVArray(i) = varSelectiveIVArray(i + 1)
                varSelectiveIVArray(i + 1) = dblTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
End Function

'Функция корректирует сортированные массивы рыночных волатильностей.
'Если в массиве объемов находятся одинаковые объемы, функция разбивает массив
'на группы, например, 10,10|9|8|7|6,6|5|2,2|3,3,3|1,1,1,1.
'Если количество групп в массиве объемов равно количеству элементов в массиве
'объемов, то происходит выход из функции, т.е. массив рыночных волатильностей
'не нуждается в корректировке.
'Иначе для каждого элемента группы определяется расстояние от центра 2-х
'центральных страйков.
'Далее происходит точечная корректирующая сортировка массива рыночных
'волатильностей в зависимости от удаленности от центра.
'Значения, которые ближе всего находятся к центру получаю более высокие места
'в массиве, а в дальнейшем и больший вес.
Private Function Correct(varSelectiveVolArray As Variant, _
                         varSelectiveStrikeArray As Variant, _
                         varSelectiveIVArray As Variant, _
                         varAvgValueCS1CS2 As Variant)
    Dim i As Integer, _
        j As Integer, _
        intCount As Integer, _
        intNumGroup As Integer, _
        intUsedItems As Integer, _
        intBegin As Integer
    Dim varArrayLength() As Variant
    'Считает количество групп.
    intCount = 1
    For i = 1 To UBound(varSelectiveVolArray) - 1
        If varSelectiveVolArray(i) <> varSelectiveVolArray(i + 1) Then
            intCount = intCount + 1
        End If
    Next i
    If intCount = UBound(varSelectiveVolArray) Then 'Проверяет наличие групп с повторяющимися элементами.
        Exit Function 'Нет необходимости в корректировке.
    Else             'Есть неоходимость в корректировке.
        intNumGroup = 1
        intUsedItems = 0
        intCount = 1
        For i = 1 To UBound(varSelectiveVolArray)
            If i > 1 Then
                If varSelectiveVolArray(i) = varSelectiveVolArray(i - 1) Then
                    intCount = intCount + 1 'Считает количество элементов в группе из повторяющихся элементов в массиве объемов.
                End If
                If varSelectiveVolArray(i) <> varSelectiveVolArray(i - 1) Or i = UBound(varSelectiveVolArray) Then 'Если точка разрыва группы или критическая точка массива.
                    intUsedItems = intUsedItems + intCount 'Фиксирует количество пройденных элементов в группах до точки разрыва.
                    'Если группа повторяющихся элементов касается верхней границы массива,
                    'т.е. начинается с 1-го элемента и принадлежит 1-й группе.
                    If intNumGroup = 1 And intCount > 1 Then
                        ReDim varArrayLength(intCount)
                        For j = 1 To UBound(varArrayLength)
                            varArrayLength(j) = Abs(varAvgValueCS1CS2 - varSelectiveStrikeArray(j))
                        Next j
                        byteTrigger = CorrectSortIV(varArrayLength, varSelectiveIVArray, intNumGroup, intBegin) 'Точечная корректирующая сортировка массива рыночных волатильностей.
                    End If
                    'Если повторяющиеся элементы где-то внутри массива, принадлежащие n-й группе.
                    If intNumGroup > 1 And intCount > 1 Then
                        ReDim varArrayLength(intUsedItems)
                        intBegin = ((intUsedItems + 1) - intCount) 'Т.О.
                        For j = intBegin To UBound(varArrayLength)
                            varArrayLength(j) = Abs(varAvgValueCS1CS2 - varSelectiveStrikeArray(j))
                        Next j
                        byteTrigger = CorrectSortIV(varArrayLength, varSelectiveIVArray, intNumGroup, intBegin) 'Точечная корректирующая сортировка массива рыночных волатильностей.
                    End If
                    intNumGroup = intNumGroup + 1 'Переход к новой группе.
                    intCount = 1 'Сброс количества элементов группы в начальное положение.
                End If
            End If
        Next i
    End If
End Function

'Функция определяет зону охвата рыночных волатильностей выше ЦС в зависимости
'от значения переменной Percentage.
Private Function lngGetPctIVArrayLengthAbvCS(varVolArray As Variant, _
                                             intNumCS2 As Integer, _
                                             varPercentage As Variant) As Long
    Dim intQtyAbvCS As Integer
    intQtyAbvCS = UBound(varVolArray) - intNumCS2
    lngGetPctIVArrayLengthAbvCS = FormatNumber((intQtyAbvCS * (varPercentage / 100)), 0) 'Округляем.
End Function

'Функция считает нулевые объемы в зоне охвата выше ЦС.
Private Function lngGetQtyZeroVolAbvCS(varVolArray As Variant, _
                                       intNumCS2 As Integer, _
                                       lngPctIVArrayLengthAbvCS As Long) As Long
    Dim i As Integer
    Dim lngQtyZeroVolAbvCS As Long
    lngQtyZeroVolAbvCS = 0
    For i = intNumCS2 + 1 To intNumCS2 + lngPctIVArrayLengthAbvCS
        If varVolArray(i) = 0 Then
            lngQtyZeroVolAbvCS = lngQtyZeroVolAbvCS + 1
        End If
    Next i
    lngGetQtyZeroVolAbvCS = lngQtyZeroVolAbvCS
End Function

'Функция определяет зону охвата рыночных волатильностей ниже ЦС в зависимости
'от значения переменной Percentage.
Private Function lngGetPctIVArrayLengthBlvCS(varVolArray As Variant, _
                                             intNumCS1 As Integer, _
                                             varPercentage As Variant) As Long
    Dim intQtyBlvCS As Integer
    intQtyBlvCS = intNumCS1 - 1
    lngGetPctIVArrayLengthBlvCS = FormatNumber((intQtyBlvCS * (varPercentage / 100)), 0) 'Округляем.
End Function

'Функция считает нулевые объемы в зоне охвата ниже ЦС.
Private Function lngGetQtyZeroVolBlvCS(varVolArray As Variant, _
                                       intNumCS1 As Integer, _
                                       lngPctIVArrayLengthBlvCS As Long) As Long
    Dim i As Integer
    Dim lngQtyZeroVolBlvCS As Long
    lngQtyZeroVolBlvCS = 0
    For i = intNumCS1 - 1 To intNumCS1 - lngPctIVArrayLengthBlvCS Step -1
        If varVolArray(i) = 0 Then
            lngQtyZeroVolBlvCS = lngQtyZeroVolBlvCS + 1
        End If
    Next i
    lngGetQtyZeroVolBlvCS = lngQtyZeroVolBlvCS
End Function
'-------------------------------------------------------------------------------
'Блок склеивания IV-начало.
'-------------------------------------------------------------------------------
'Priority(max->min):varIVCSWithBigVol->varIVCSWithSmallVol->
'varSelectiveIVArray->varIVArrayAbvCS->varIVArrayBlvCS
'(varIVArrayAbvCS>varIVArrayBlvCS, учитывается логнормальное распределение).
'Var1.Функция выполняет склеивание массивов рыночных волатильностей,
'если нулевых объемов выше и ниже ЦС нет.
Private Function varGluing1(lngQtyNonZeroVol As Long, _
                            lngTotalSizeIVArray As Long, _
                            varIVCSWithBigVol As Variant, _
                            varIVCSWithSmallVol As Variant, _
                            varSelectiveIVArray As Variant) As Variant
    Dim i As Integer, _
        varTotalIVArray() As Variant
    lngTotalSizeIVArray = 2 + lngQtyNonZeroVol 'Определяем общий массив IV.
    ReDim varTotalIVArray(lngTotalSizeIVArray)
    For i = 1 To lngTotalSizeIVArray
        If i = 1 Then
            varTotalIVArray(i) = varIVCSWithBigVol
        ElseIf i = 2 Then
            varTotalIVArray(i) = varIVCSWithSmallVol
        ElseIf i > 2 And i <= lngTotalSizeIVArray Then
            varTotalIVArray(i) = varSelectiveIVArray(i - 2) 'Получаем общий массив IV.
        End If
    Next i
    varGluing1 = varTotalIVArray
End Function

'Var2.Функция выполняет склеивание массивов рыночных волатильностей,
'если нулевые объемы выше ЦС есть, ниже ЦС нет.
Private Function varGluing2(intNumCS2 As Integer, _
                            lngPctIVArrayLengthAbvCS As Long, _
                            varVolArray As Variant, _
                            varIVArray As Variant, _
                            lngQtyNonZeroVol As Long, _
                            lngQtyZeroVolAbvCS As Long, _
                            lngTotalSizeIVArray As Long, _
                            varIVCSWithBigVol As Variant, _
                            varIVCSWithSmallVol As Variant, _
                            varSelectiveIVArray As Variant) As Variant
    Dim i As Integer, _
        intCount As Integer, _
        varIVArrayAbvCS() As Variant, _
        varTotalIVArray() As Variant
    ReDim varIVArrayAbvCS(lngQtyZeroVolAbvCS) 'Определяем массив IV выше ЦС.
    intCount = 0
    For i = intNumCS2 + 1 To intNumCS2 + lngPctIVArrayLengthAbvCS
        If varVolArray(i) = 0 Then
            intCount = intCount + 1
            varIVArrayAbvCS(intCount) = varIVArray(i) 'Получаем массив IV выше ЦС.
        End If
    Next i
    lngTotalSizeIVArray = 2 + lngQtyNonZeroVol + lngQtyZeroVolAbvCS 'Определяем общий массив IV.
    ReDim varTotalIVArray(lngTotalSizeIVArray)
    For i = 1 To lngTotalSizeIVArray 'Получаем общий массив IV.
        If i = 1 Then
            varTotalIVArray(i) = varIVCSWithBigVol
        ElseIf i = 2 Then
            varTotalIVArray(i) = varIVCSWithSmallVol
        ElseIf i > 2 And i <= (lngQtyNonZeroVol + 2) Then
            varTotalIVArray(i) = varSelectiveIVArray(i - 2)
        ElseIf i > (lngQtyNonZeroVol + 2) And i <= lngTotalSizeIVArray Then
            varTotalIVArray(i) = varIVArrayAbvCS(i - (lngQtyNonZeroVol + 2))
        End If
    Next i
    varGluing2 = varTotalIVArray
End Function

'Var3.Функция выполняет склеивание массивов рыночных волатильностей,
'если нулевых объемов выше ЦС нет, ниже ЦС есть.
Private Function varGluing3(intNumCS1 As Integer, _
                            lngPctIVArrayLengthBlvCS As Long, _
                            varVolArray As Variant, _
                            varIVArray As Variant, _
                            lngQtyNonZeroVol As Long, _
                            lngQtyZeroVolBlvCS As Long, _
                            lngTotalSizeIVArray As Long, _
                            varIVCSWithBigVol As Variant, _
                            varIVCSWithSmallVol As Variant, _
                            varSelectiveIVArray As Variant) As Variant
    Dim i As Integer, _
        intCount As Integer, _
        varIVArrayBlvCS() As Variant, _
        varTotalIVArray() As Variant
    ReDim varIVArrayBlvCS(lngQtyZeroVolBlvCS) 'Определяем массив IV ниже ЦС.
    intCount = 0
    For i = intNumCS1 - 1 To intNumCS1 - lngPctIVArrayLengthBlvCS Step -1
        If varVolArray(i) = 0 Then
            intCount = intCount + 1
            varIVArrayBlvCS(intCount) = varIVArray(i) 'Получаем массив IV ниже ЦС.
        End If
    Next i
    lngTotalSizeIVArray = 2 + lngQtyNonZeroVol + lngQtyZeroVolBlvCS 'Определяем общий массив IV.
    ReDim varTotalIVArray(lngTotalSizeIVArray)
    For i = 1 To lngTotalSizeIVArray 'Получаем общий массив IV.
        If i = 1 Then
            varTotalIVArray(i) = varIVCSWithBigVol
        ElseIf i = 2 Then
            varTotalIVArray(i) = varIVCSWithSmallVol
        ElseIf i > 2 And i <= (lngQtyNonZeroVol + 2) Then
            varTotalIVArray(i) = varSelectiveIVArray(i - 2)
        ElseIf i > (lngQtyNonZeroVol + 2) And i <= lngTotalSizeIVArray Then
            varTotalIVArray(i) = varIVArrayBlvCS(i - (lngQtyNonZeroVol + 2))
        End If
    Next i
    varGluing3 = varTotalIVArray
End Function

'Var4.Функция выполняет склеивание массивов рыночных волатильностей,
'если нулевые объемы выше и ниже ЦС есть.
Private Function varGluing4(intNumCS2 As Integer, _
                            intNumCS1 As Integer, _
                            lngPctIVArrayLengthAbvCS As Long, _
                            lngPctIVArrayLengthBlvCS As Long, _
                            varVolArray As Variant, _
                            varIVArray As Variant, _
                            lngQtyNonZeroVol As Long, _
                            lngQtyZeroVolAbvCS As Long, _
                            lngQtyZeroVolBlvCS As Long, _
                            lngTotalSizeIVArray As Long, _
                            varIVCSWithBigVol As Variant, _
                            varIVCSWithSmallVol As Variant, _
                            varSelectiveIVArray As Variant) As Variant
    Dim i As Integer, _
        intCount As Integer, _
        varIVArrayAbvCS() As Variant, _
        varIVArrayBlvCS() As Variant, _
        varTotalIVArray() As Variant
    ReDim varIVArrayAbvCS(lngQtyZeroVolAbvCS) 'Определяем массив IV выше ЦС.
    intCount = 0
    For i = intNumCS2 + 1 To intNumCS2 + lngPctIVArrayLengthAbvCS
        If varVolArray(i) = 0 Then
            intCount = intCount + 1
            varIVArrayAbvCS(intCount) = varIVArray(i) 'Получаем массив IV выше ЦС.
        End If
    Next i
    ReDim varIVArrayBlvCS(lngQtyZeroVolBlvCS) 'Определяем массив IV ниже ЦС.
    intCount = 0
    For i = intNumCS1 - 1 To intNumCS1 - lngPctIVArrayLengthBlvCS Step -1
        If varVolArray(i) = 0 Then
            intCount = intCount + 1
            varIVArrayBlvCS(intCount) = varIVArray(i) 'Получаем массив IV ниже ЦС.
        End If
    Next i
    lngTotalSizeIVArray = 2 + lngQtyNonZeroVol + lngQtyZeroVolAbvCS + lngQtyZeroVolBlvCS 'Определяем общий массив IV.
    ReDim varTotalIVArray(lngTotalSizeIVArray)
    For i = 1 To lngTotalSizeIVArray 'Получаем общий массив IV.
        If i = 1 Then
            varTotalIVArray(i) = varIVCSWithBigVol
        ElseIf i = 2 Then
            varTotalIVArray(i) = varIVCSWithSmallVol
        ElseIf i > 2 And i <= (lngQtyNonZeroVol + 2) Then
            varTotalIVArray(i) = varSelectiveIVArray(i - 2)
        ElseIf i > (lngQtyNonZeroVol + 2) And i <= (lngQtyZeroVolAbvCS + lngQtyNonZeroVol + 2) Then
            varTotalIVArray(i) = varIVArrayAbvCS(i - (lngQtyNonZeroVol + 2))
        ElseIf i > (lngQtyZeroVolAbvCS + lngQtyNonZeroVol + 2) And i <= lngTotalSizeIVArray Then
            varTotalIVArray(i) = varIVArrayBlvCS(i - (lngQtyZeroVolAbvCS + lngQtyNonZeroVol + 2))
        End If
    Next i
    varGluing4 = varTotalIVArray
End Function

'-------------------------------------------------------------------------------
'Блок склеивания IV-конец.
'-------------------------------------------------------------------------------
'Функция создает массив весов, используя формулу Фишберна (2*(n-i+1))/(n*(n+1)),
'где n - количество элементов в массиве, i - номер элемента в массиве.
Private Function varGetWeightedCoeffFishburn(lngTotalSizeIVArray As Long) As Variant
    Dim i As Integer
    Dim varWCF() As Variant, _
        varWeight As Variant
    ReDim varWCF(lngTotalSizeIVArray)
    For i = 1 To lngTotalSizeIVArray
        varWeight = (2 * (lngTotalSizeIVArray - i + 1)) / (lngTotalSizeIVArray * (lngTotalSizeIVArray + 1))
        varWCF(i) = varWeight
    Next i
    varGetWeightedCoeffFishburn = varWCF
End Function

'Функция считает итоговую сумму значений взвешенных с помощью коэффициентов
'Фишберна (агрегированное значение).
Private Function varGetTotalWeightSumIV(varTotalIVArray As Variant, _
                                        varWeightedCoeffFishburn As Variant) As Variant
    Dim varWeightSumIV As Variant, _
        varWeightPartIV As Variant
    Dim i As Integer
    varWeightSumIV = 0
    For i = 1 To UBound(varTotalIVArray)
        varWeightPartIV = varWeightedCoeffFishburn(i) * varTotalIVArray(i)
        varWeightSumIV = varWeightSumIV + varWeightPartIV
    Next i
    varGetTotalWeightSumIV = varWeightSumIV
End Function

'Функция дает IV.
Private Function varAggIV(varFuturesPrice As Variant, _
                          varPercentage As Variant, _
                          varStrikeArray As Variant, _
                          varIVArray As Variant, _
                          varVolArray As Variant) As Variant

    Dim intNumCS1 As Integer, _
        intNumCS2 As Integer, _
        varIVCSWithBigVol As Variant, _
        varIVCSWithSmallVol As Variant, _
        varAvgValueCS1CS2 As Variant

    Dim lngQtyNonZeroVol As Long, _
        varSelectiveVolArray() As Variant, _
        varSelectiveStrikeArray() As Variant, _
        varSelectiveIVArray() As Variant

    Dim lngPctIVArrayLengthAbvCS As Long, _
        lngPctIVArrayLengthBlvCS As Long, _
        lngQtyZeroVolAbvCS As Long, _
        lngQtyZeroVolBlvCS As Long, _
        lngTotalSizeIVArray As Long

    Dim varGluing As Variant, _
        varWeightedCoeffFishburn As Variant

    'Получаем значения intNumCS1, intNumCS2, varIVCSWithBigVol, varIVCSWithSmallVol
    'т.е. записываем номера 2-х центральных страйков и соответствующие им
    'волатильности, в зависимости от объемов на этих страйках.
    byteTrigger = CentralStrike(varFuturesPrice, _
                                varVolArray, _
                                varStrikeArray, _
                                varIVArray, _
                                intNumCS1, _
                                intNumCS2, _
                                varIVCSWithBigVol, _
                                varIVCSWithSmallVol, _
                                varAvgValueCS1CS2)
    'Считаем количество всех ненулевых объемов, не затрагивая значения 2-х
    'центральных страйков.
    lngQtyNonZeroVol = lngGetQtyNonZeroVol(varVolArray, _
                                           intNumCS1, _
                                           intNumCS2, _
                                           lngQtyNonZeroVol)
    'Если все объемы равны нулю, то функция возвращает значение волатильности
    'на центральном страйке.
    If lngQtyNonZeroVol = 0 Then
        If varFuturesPrice >= varStrikeArray(intNumCS1) And varFuturesPrice < varAvgValueCS1CS2 Then
            varAggIV = 0.8 * varIVArray(intNumCS1) + 0.2 * varIVArray(intNumCS2)
        ElseIf varFuturesPrice = varAvgValueCS1CS2 Then
            varAggIV = 0.5 * varIVArray(intNumCS1) + 0.5 * varIVArray(intNumCS2)
        ElseIf varFuturesPrice > varAvgValueCS1CS2 And varFuturesPrice <= varStrikeArray(intNumCS2) Then
            varAggIV = 0.2 * varIVArray(intNumCS1) + 0.8 * varIVArray(intNumCS2)
        End If
    Else
        'Выбираем все ненулевые объемы, соответствующие им волатильности
        'и страйки, не затрагивая значения 2-х центральных страйков, и сохраняем
        'в массивах (varSelectiveVolArray, varSelectiveStrikeArray, varSelectiveIVArray).
        ReDim varSelectiveVolArray(lngQtyNonZeroVol)
        ReDim varSelectiveStrikeArray(lngQtyNonZeroVol)
        ReDim varSelectiveIVArray(lngQtyNonZeroVol)
        byteTrigger = Selection(varVolArray, _
                                varStrikeArray, _
                                varIVArray, _
                                intNumCS1, _
                                intNumCS2, _
                                varSelectiveVolArray, _
                                varSelectiveStrikeArray, _
                                varSelectiveIVArray)
        'Сортируем массивы varSelectiveVolArray, varSelectiveStrikeArray
        'и varSelectiveIVArray.
        byteTrigger = SortIV(varSelectiveVolArray, _
                             varSelectiveStrikeArray, _
                             varSelectiveIVArray)
        'Корректируем сортированные массивы рыночных волатильностей.
        byteTrigger = Correct(varSelectiveVolArray, _
                              varSelectiveStrikeArray, _
                              varSelectiveIVArray, _
                              varAvgValueCS1CS2)
        'Определяем зону охвата рыночных волатильностей выше ЦС в зависимости
        'от значения переменной Percentage.
        lngPctIVArrayLengthAbvCS = lngGetPctIVArrayLengthAbvCS(varVolArray, _
                                                               intNumCS2, _
                                                               varPercentage)
        'Считаем нулевые объемы в зоне охвата выше ЦС.
        lngQtyZeroVolAbvCS = lngGetQtyZeroVolAbvCS(varVolArray, _
                                                   intNumCS2, _
                                                   lngPctIVArrayLengthAbvCS)
        'Определяем зону охвата рыночных волатильностей ниже ЦС в зависимости
        'от значения переменной Percentage.
        lngPctIVArrayLengthBlvCS = lngGetPctIVArrayLengthBlvCS(varVolArray, _
                                                               intNumCS1, _
                                                               varPercentage)
        'Считаем нулевые объемы в зоне охвата ниже ЦС.
        lngQtyZeroVolBlvCS = lngGetQtyZeroVolBlvCS(varVolArray, _
                                                   intNumCS1, _
                                                   lngPctIVArrayLengthBlvCS)
        'Склеиваем массивы волатильностей.
        If lngQtyZeroVolAbvCS = 0 And lngQtyZeroVolBlvCS = 0 Then 'Var1. Если нулевых объемов выше и ниже ЦС нет.
            varGluing = varGluing1(lngQtyNonZeroVol, _
                                   lngTotalSizeIVArray, _
                                   varIVCSWithBigVol, _
                                   varIVCSWithSmallVol, _
                                   varSelectiveIVArray)
        ElseIf lngQtyZeroVolAbvCS > 0 And lngQtyZeroVolBlvCS = 0 Then 'Var2. Если нулевые объемы выше ЦС есть, ниже ЦС нет.
            varGluing = varGluing2(intNumCS2, _
                                   lngPctIVArrayLengthAbvCS, _
                                   varVolArray, _
                                   varIVArray, _
                                   lngQtyNonZeroVol, _
                                   lngQtyZeroVolAbvCS, _
                                   lngTotalSizeIVArray, _
                                   varIVCSWithBigVol, _
                                   varIVCSWithSmallVol, _
                                   varSelectiveIVArray)
        ElseIf lngQtyZeroVolAbvCS = 0 And lngQtyZeroVolBlvCS > 0 Then 'Var3. Если нулевых объемов выше ЦС нет, ниже ЦС есть.
            varGluing = varGluing3(intNumCS1, _
                                   lngPctIVArrayLengthBlvCS, _
                                   varVolArray, _
                                   varIVArray, _
                                   lngQtyNonZeroVol, _
                                   lngQtyZeroVolBlvCS, _
                                   lngTotalSizeIVArray, _
                                   varIVCSWithBigVol, _
                                   varIVCSWithSmallVol, _
                                   varSelectiveIVArray)
        ElseIf lngQtyZeroVolAbvCS > 0 And lngQtyZeroVolBlvCS > 0 Then 'Var4. Если нулевые объемы выше и ниже ЦС есть.
            varGluing = varGluing4(intNumCS2, _
                                   intNumCS1, _
                                   lngPctIVArrayLengthAbvCS, _
                                   lngPctIVArrayLengthBlvCS, _
                                   varVolArray, _
                                   varIVArray, _
                                   lngQtyNonZeroVol, _
                                   lngQtyZeroVolAbvCS, _
                                   lngQtyZeroVolBlvCS, _
                                   lngTotalSizeIVArray, _
                                   varIVCSWithBigVol, _
                                   varIVCSWithSmallVol, _
                                   varSelectiveIVArray)
        End If
        'Создаем массив весов.
        varWeightedCoeffFishburn = varGetWeightedCoeffFishburn(lngTotalSizeIVArray)
        'Выполняем итоговое взвешивание и суммирование волатильностей.
        varAggIV = varGetTotalWeightSumIV(varGluing, varWeightedCoeffFishburn)
    End If
End Function

'-------------------------------------------------------------------------------
'Блок IV-конец.
'-------------------------------------------------------------------------------
'Функция выводит отчетные даты.
Private Function OutputDateArrayFromReportFileNames(varDateArrayFromReportFileNames As Variant, _
                                                    intRow_TradeDate As Integer, _
                                                    intColumn_TradeDate As Integer)
    Dim i As Integer
    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("AutoIV").Range(Cells(intRow_TradeDate + 2, intColumn_TradeDate), Cells(intRow_TradeDate + 300, intColumn_TradeDate)).ClearContents
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        Worksheets("AutoIV").Cells(intRow_TradeDate + 1 + i, intColumn_TradeDate).Value = varDateArrayFromReportFileNames(i)
    Next i
End Function

'Функция выводит агрегированные IV.
Private Function OutputAggIV(varDateArrayFromReportFileNames As Variant, _
                             varAggIV As Variant, _
                             intRow_AggIV As Integer, _
                             intColumn_AggIV As Integer)
    Dim i As Integer
    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("AutoIV").Range(Cells(intRow_AggIV + 2, intColumn_AggIV), Cells(intRow_AggIV + 300, intColumn_AggIV)).ClearContents
    For i = 1 To UBound(varDateArrayFromReportFileNames)
        Worksheets("AutoIV").Cells(intRow_AggIV + 1 + i, intColumn_AggIV).Value = varAggIV(i)
    Next i
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Название и формат полей CSV-файла pr-options-delta-ru.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Дата      |Код опциона                   |Право  |Последний торговый день,     |Цена исполнения|Кол-во подлежащих   |Сумма подлежащих    |Коэф. |Теор.  |Вола.|
'          |                              |по     |в который может быть         |опционного     |исполнению опционных|исполнению опционных|дельта|цена   |     |
'          |                              |опциону|заключен опционный контракт  |контракта      |контрактов          |контрактов          |      |опциона|     |
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'dd.mm.yyyy;КодБА-mm.yyMddmmyyCAЦенаСтрайк;Call   ;yyyy-mm-dd                   ;               ;                    ;                    ;      ;       ;     ;
'dd.mm.yyyy;КодБА-mm.yyMddmmyyPAЦенаСтрайк;Put    ;yyyy-mm-dd                   ;               ;                    ;                    ;      ;       ;     ;
'          1                              2       3                             4               5                    6                    7      8       9     10
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Название и формат полей CSV-файла rts-results-ru.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Код контракта                 |Дата      |Средневзвешенная|Расчетная|Цена первой|Максимальная|Минимальная|Цена     |Изменение|Объем    |Число |Объем  |Объем  |Объем   |Объем   |
'                              |исполнения|цена за лот     |цена     |сделки     |цена        |цена       |последней|         |последней|сделок|торгов,|торгов,|открытых|открытых|
'                              |          |                |         |           |            |           |сделки   |         |сделки,  |      |руб.   |контр. |позиций,|позиций,|
'                              |          |                |         |           |            |           |         |         |контр.   |      |       |       |руб.    |кол-во  |
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'КодБА-mm.yyMddmmyyCAЦенаСтрайк;dd.mm.yyyy;                ;         ;           ;            ;           ;         ;         ;         ;      ;       ;       ;        ;        ;
'КодБА-mm.yyMddmmyyPAЦенаСтрайк;dd.mm.yyyy;                ;         ;           ;            ;           ;         ;         ;         ;      ;       ;       ;        ;        ;
'                              1          2                3         4           5            6           7         8         9         10     11      12      13       14       15
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub AutoIV()
Attribute AutoIV.VB_Description = "Дает агрегированное значение IV."
Attribute AutoIV.VB_ProcData.VB_Invoke_Func = "i\n14"
    Const SBRF_Name As String = "SBRF"
    Const Si_Name As String = "Si"
    Const RI_Name As String = "RTS"
    Const BR_Name As String = "BR"
    Const GD_Name As String = "GOLD"

    Dim i As Integer, j As Integer

    Dim intRow_Ssf_MO_ExpDate As Integer, intColumn_Ssf_MO_ExpDate As Integer, _
        intRow_Fxf_MO_ExpDate As Integer, intColumn_Fxf_MO_ExpDate As Integer, _
        intRow_Inf_MO_ExpDate As Integer, intColumn_Inf_MO_ExpDate As Integer, _
        intRow_Enf_MO_ExpDate As Integer, intColumn_Enf_MO_ExpDate As Integer, _
        intRow_Pmf_QO_ExpDate As Integer, intColumn_Pmf_QO_ExpDate As Integer, _
        intRow_Ssf_WO_ExpDate As Integer, intColumn_Ssf_WO_ExpDate As Integer, _
        intRow_InfFxf_WO_ExpDate As Integer, intColumn_InfFxf_WO_ExpDate As Integer, _
        intRow_Enf_WO_ExpDate As Integer, intColumn_Enf_WO_ExpDate As Integer

    Dim intRow_Directory As Integer, intColumn_Directory As Integer, _
        intRow_ProductCode As Integer, intColumn_ProductCode As Integer, _
        intRow_Lifespan As Integer, intColumn_Lifespan As Integer, _
        intRow_Combo As Integer, intColumn_Combo As Integer, _
        intRow_Percentage As Integer, intColumn_Percentage As Integer

    Dim intRow_TradeDate As Integer, intColumn_TradeDate As Integer, _
        intRow_FrontAggIVCall As Integer, intColumn_FrontAggIVCall As Integer, _
        intRow_FrontAggIVPut As Integer, intColumn_FrontAggIVPut As Integer, _
        intRow_FrontAvgAggIV As Integer, intColumn_FrontAvgAggIV As Integer, _
        intRow_NextAggIVCall As Integer, intColumn_NextAggIVCall As Integer, _
        intRow_NextAggIVPut As Integer, intColumn_NextAggIVPut As Integer, _
        intRow_NextAvgAggIV As Integer, intColumn_NextAvgAggIV As Integer, _
        intRow_BackAggIVCall As Integer, intColumn_BackAggIVCall As Integer, _
        intRow_BackAggIVPut As Integer, intColumn_BackAggIVPut As Integer, _
        intRow_BackAvgAggIV As Integer, intColumn_BackAvgAggIV As Integer, _
        intRow_ComboAvgAggIV As Integer, intColumn_ComboAvgAggIV As Integer

    Dim varDirectory As Variant, _
        varProductCode As Variant, _
        varLifespan As Variant, _
        varCombo As Variant, _
        varPercentage As Variant, _
        varExpDateArray As Variant

    Dim intQtyReportFiles As Integer, _
        intQtyDatesFromReportFileNames As Integer, _
        varReportFileNameArray As Variant, _
        varDateArrayFromReportFileNames As Variant, _
        varReportFileNameArray_pr_options_delta_ru As Variant, _
        varReportFileNameArray_rts_results_ru As Variant, _
        boolCheckDualDates As Boolean, _
        varFuturesPriceArray As Variant

    Dim varFrontExpDateArray() As Variant, _
        varNextExpDateArray() As Variant, _
        varBackExpDateArray() As Variant

    Dim varCallStrikeArray() As Variant, _
        varPutStrikeArray() As Variant, _
        varCallIVArray() As Variant, _
        varPutIVArray() As Variant, _
        varCallVolArray() As Variant, _
        varPutVolArray() As Variant, _
        boolIsThereExpDate_pr_options_delta_ru As Boolean, _
        boolIsThereExpDate_rts_results_ru As Boolean

    Dim varFrontAggIVCall() As Variant, _
        varFrontAggIVPut() As Variant, _
        varFrontAvgAggIV() As Variant, _
        varNextAggIVCall() As Variant, _
        varNextAggIVPut() As Variant, _
        varNextAvgAggIV() As Variant, _
        varBackAggIVCall() As Variant, _
        varBackAggIVPut() As Variant, _
        varBackAvgAggIV() As Variant, _
        varComboAvgAggIV() As Variant
    'Получаем адреса ячеек входных параметров с панели пресетов.
    byteTrigger = GetCellAddress("Ecal", intRow_Ssf_MO_ExpDate, intColumn_Ssf_MO_ExpDate, "SSF-MO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_Fxf_MO_ExpDate, intColumn_Fxf_MO_ExpDate, "FXF-MO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_Inf_MO_ExpDate, intColumn_Inf_MO_ExpDate, "INF-MO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_Enf_MO_ExpDate, intColumn_Enf_MO_ExpDate, "ENF-MO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_Pmf_QO_ExpDate, intColumn_Pmf_QO_ExpDate, "PMF-QO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_Ssf_WO_ExpDate, intColumn_Ssf_WO_ExpDate, "SSF-WO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_InfFxf_WO_ExpDate, intColumn_InfFxf_WO_ExpDate, "INF&FXF-WO-ExpDate")
    byteTrigger = GetCellAddress("Ecal", intRow_Enf_WO_ExpDate, intColumn_Enf_WO_ExpDate, "ENF-WO-ExpDate")
    byteTrigger = GetCellAddress("AutoIV", intRow_Directory, intColumn_Directory, "Directory")
    byteTrigger = GetCellAddress("AutoIV", intRow_ProductCode, intColumn_ProductCode, "ProdCode")
    byteTrigger = GetCellAddress("AutoIV", intRow_Lifespan, intColumn_Lifespan, "Lifespan")
    byteTrigger = GetCellAddress("AutoIV", intRow_Combo, intColumn_Combo, "Combo")
    byteTrigger = GetCellAddress("AutoIV", intRow_Percentage, intColumn_Percentage, "Percentage")
    byteTrigger = GetCellAddress("AutoIV", intRow_TradeDate, intColumn_TradeDate, "TradeDate")
    byteTrigger = GetCellAddress("AutoIV", intRow_FrontAggIVCall, intColumn_FrontAggIVCall, "FrontAggIV_Call")
    byteTrigger = GetCellAddress("AutoIV", intRow_FrontAggIVPut, intColumn_FrontAggIVPut, "FrontAggIV_Put")
    byteTrigger = GetCellAddress("AutoIV", intRow_FrontAvgAggIV, intColumn_FrontAvgAggIV, "FrontAvgAggIV")
    byteTrigger = GetCellAddress("AutoIV", intRow_NextAggIVCall, intColumn_NextAggIVCall, "NextAggIV_Call")
    byteTrigger = GetCellAddress("AutoIV", intRow_NextAggIVPut, intColumn_NextAggIVPut, "NextAggIV_Put")
    byteTrigger = GetCellAddress("AutoIV", intRow_NextAvgAggIV, intColumn_NextAvgAggIV, "NextAvgAggIV")
    byteTrigger = GetCellAddress("AutoIV", intRow_BackAggIVCall, intColumn_BackAggIVCall, "BackAggIV_Call")
    byteTrigger = GetCellAddress("AutoIV", intRow_BackAggIVPut, intColumn_BackAggIVPut, "BackAggIV_Put")
    byteTrigger = GetCellAddress("AutoIV", intRow_BackAvgAggIV, intColumn_BackAvgAggIV, "BackAvgAggIV")
    byteTrigger = GetCellAddress("AutoIV", intRow_ComboAvgAggIV, intColumn_ComboAvgAggIV, "ComboAvgAggIV")
    'Считываем значения с панели пресетов.
    varDirectory = Worksheets("AutoIV").Cells(intRow_Directory, intColumn_Directory + 1).Value
    varProductCode = Worksheets("AutoIV").Cells(intRow_ProductCode, intColumn_ProductCode + 1).Value
    varLifespan = Worksheets("AutoIV").Cells(intRow_Lifespan, intColumn_Lifespan + 1).Value
    varCombo = Worksheets("AutoIV").Cells(intRow_Combo, intColumn_Combo + 1).Value
    varPercentage = Worksheets("AutoIV").Cells(intRow_Percentage, intColumn_Percentage + 1).Value
    If varCombo = Empty Then 'Значение по умолчанию, если параметр Combo пуст.
        varCombo = 1
    End If
    If varPercentage = Empty Then 'Значение по умолчанию, если параметр Percentage пуст.
        varPercentage = 0
    End If
    If varProductCode = Empty Then 'Условие для полной очистки рабочего пространства.
        Worksheets("AutoIV").Range(Cells(intRow_TradeDate + 2, intColumn_TradeDate), _
        Cells(intRow_ComboAvgAggIV + 300, intColumn_ComboAvgAggIV)).ClearContents
        'Проверяем директорию.
    ElseIf varDirectory = Empty Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Не указан путь к папке с файл-отчетами." _
        & vbNewLine & "Добавьте значение параметра Directory и повторите расчет IV.")
        'Проверяем торговый инструмент.
    ElseIf UCase(varProductCode) <> "SBRF" And UCase(varProductCode) <> "SR" And _
        UCase(varProductCode) <> "SI" And _
        UCase(varProductCode) <> "RTS" And UCase(varProductCode) <> "RI" And _
        UCase(varProductCode) <> "BRENT" And UCase(varProductCode) <> "BR" And _
        UCase(varProductCode) <> "GOLD" And UCase(varProductCode) <> "GD" Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Текущий торговый инструмент не поддерживается." _
        & vbNewLine & "Список доступных инструментов:" _
        & vbNewLine & "SBRF;" _
        & vbNewLine & "Si;" _
        & vbNewLine & "RI;" _
        & vbNewLine & "BR;" _
        & vbNewLine & "GD." _
        & vbNewLine & "Измените значение параметра ProductCode и повторите расчет IV.")
        'Проверяем срок обращения.
    ElseIf UCase(varLifespan) <> "W" And UCase(varLifespan) <> "M" And UCase(varLifespan) <> "Q" Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Текущий срок обращения не поддерживается." _
        & vbNewLine & "Список доступных сроков обращения:" _
        & vbNewLine & "W (недельный);" _
        & vbNewLine & "M (месячный);" _
        & vbNewLine & "Q (квартальный)." _
        & vbNewLine & "Измените значение параметра Lifespan и повторите расчет IV.")
        'Проверяем комбинации.
    ElseIf varCombo <> 1 And varCombo <> 2 And varCombo <> 3 Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Текущая комбинация не поддерживается." _
        & vbNewLine & "Список доступных комбинаций:" _
        & vbNewLine & "1 (ближайшая серия);" _
        & vbNewLine & "2 (ближайшая серия+следующая за ближайшей);" _
        & vbNewLine & "3 (ближайшая серия+следующая за ближайшей+следующая за следующей)." _
        & vbNewLine & "NB! Если параметр Combo оставить пустым, то по умолчанию его значение будет 1." _
        & vbNewLine & "Измените значение параметра Combo и повторите расчет IV.")
        'Проверяем процент охвата.
    ElseIf varPercentage < 0 Or varPercentage > 100 Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Текущее значение охвата находится за пределом допустимых значений." _
        & vbNewLine & "Диапазон допустимых значений:" _
        & vbNewLine & "Любые положительные числа от 0 до 100" _
        & vbNewLine & "NB! Если параметр Percentage оставить пустым, то по умолчанию его значение будет 0." _
        & vbNewLine & "Измените значение параметра Percentage и повторите расчет IV.")
    Else
        'Создаем соответствие регистру имен файл-отчетов.
        If UCase(varProductCode) = "SBRF" Then
            varProductCode = SBRF_Name
        End If
        If UCase(varProductCode) = "SI" Then
            varProductCode = Si_Name
        End If
        If UCase(varProductCode) = "RI" Then
            varProductCode = RI_Name
        End If
        If UCase(varProductCode) = "BR" Then
            varProductCode = BR_Name
        End If
        If UCase(varProductCode) = "GD" Then
            varProductCode = GD_Name
        End If
        'Считываем даты экспирации.
        If varProductCode = "SBRF" Then
            If UCase(varLifespan) = "Q" Or UCase(varLifespan) = "M" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Ssf_MO_ExpDate, intColumn_Ssf_MO_ExpDate, varLifespan)
            ElseIf UCase(varLifespan) = "W" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Ssf_WO_ExpDate, intColumn_Ssf_WO_ExpDate, varLifespan)
            End If
        ElseIf varProductCode = "Si" Then
            If UCase(varLifespan) = "Q" Or UCase(varLifespan) = "M" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Fxf_MO_ExpDate, intColumn_Fxf_MO_ExpDate, varLifespan)
            ElseIf UCase(varLifespan) = "W" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_InfFxf_WO_ExpDate, intColumn_InfFxf_WO_ExpDate, varLifespan)
            End If
        ElseIf varProductCode = "RTS" Then
            If UCase(varLifespan) = "Q" Or UCase(varLifespan) = "M" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Inf_MO_ExpDate, intColumn_Inf_MO_ExpDate, varLifespan)
            ElseIf UCase(varLifespan) = "W" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_InfFxf_WO_ExpDate, intColumn_InfFxf_WO_ExpDate, varLifespan)
            End If
        ElseIf varProductCode = "BR" Then
            If UCase(varLifespan) = "Q" Then
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Квартальных опционов на нефть нет." _
                & vbNewLine & "Измените значение параметра Lifespan на M и повторите расчет IV.")
            ElseIf UCase(varLifespan) = "M" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Enf_MO_ExpDate, intColumn_Enf_MO_ExpDate, varLifespan)
            ElseIf UCase(varLifespan) = "W" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Enf_WO_ExpDate, intColumn_Enf_WO_ExpDate, varLifespan)
            End If
        ElseIf varProductCode = "GOLD" Then
            If UCase(varLifespan) = "Q" Then
                varExpDateArray = varGetExpDateArray("Ecal", intRow_Pmf_QO_ExpDate, intColumn_Pmf_QO_ExpDate, varLifespan)
            ElseIf UCase(varLifespan) = "M" Then
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Месячных опционов на золото нет." _
                & vbNewLine & "Измените значение параметра Lifespan на Q и повторите расчет IV.")
            ElseIf UCase(varLifespan) = "W" Then
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Недельных опционов на золото нет." _
                & vbNewLine & "Измените значение параметра Lifespan на Q и повторите расчет IV.")
            End If
        End If
        If IsEmpty(varExpDateArray) = False Then 'Если массив дат экспирации не пуст.
            intQtyReportFiles = intGetQtyReportFiles(varDirectory) 'Считаем количество файл-отчетов в папке назначения.
            If intQtyReportFiles > 0 Then 'Если в папке есть файл-отчеты.
                varReportFileNameArray = varGetReportFileNameArray(varDirectory, intQtyReportFiles) 'Создаем массив имен файл-отчетов.
                varDateArrayFromReportFileNames = varGetDateArrayFromReportFileNames(varReportFileNameArray) 'Создаем массив дат файл-отчетов.
                intQtyDatesFromReportFileNames = intGetQtyDatesFromReportFileNames(varDateArrayFromReportFileNames) 'Считаем количество дат.
                'I. Если количество дат файл-отчетов pr-options-delta-ru и rts-results-ru не совпадает с количеством файл-отчетов pr-options-delta-ru и rts-results-ru.
                If intQtyDatesFromReportFileNames <> intQtyReportFiles Then
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "В каком-то файл-отчете pr-options-delta-ru или rts-results-ru нет даты." _
                    & vbNewLine & "Проверьте файл-отчеты в папке назначения и повторите расчет IV.")
                Else
                    byteTrigger = BubbleSort(varReportFileNameArray, varDateArrayFromReportFileNames) 'Сортируем.
                    varReportFileNameArray_pr_options_delta_ru = Filter(varReportFileNameArray, "pr-options-delta-ru") 'Сюда фильтруем файл-отчеты pr-options-delta-ru.
                    varReportFileNameArray_pr_options_delta_ru = varToOptionBase1(varReportFileNameArray_pr_options_delta_ru) 'Приводим массив в соответствие с оператором Option Base 1.
                    varReportFileNameArray_rts_results_ru = Filter(varReportFileNameArray, "rts-results-ru") 'Сюда фильтруем файл-отчеты rts-results-ru.
                    varReportFileNameArray_rts_results_ru = varToOptionBase1(varReportFileNameArray_rts_results_ru) 'Приводим массив в соответствие с оператором Option Base 1.
                    varReportFileNameArray = Empty
                    'II. Если количество файл-отчетов pr-options-delta-ru и rts-results-ru не совпадает.
                    If UBound(varReportFileNameArray_pr_options_delta_ru) <> UBound(varReportFileNameArray_rts_results_ru) Then
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "Количество файл-отчетов pr-options-delta-ru и rts-results-ru должно быть одинаковое." _
                        & vbNewLine & "Проверьте файл-отчеты в папке назначения и повторите расчет IV.")
                    Else
                        'III. Если пара файл-отчетов pr-options-delta-ru и rts-results-ru имеет разные даты.
                        boolCheckDualDates = boolGetCheckDualDates(varDateArrayFromReportFileNames)
                        If boolCheckDualDates = False Then
                            MsgBox ("Оба-на!!!" _
                            & vbNewLine & "Пара файл-отчетов pr-options-delta-ru и rts-results-ru должна иметь одинаковые даты." _
                            & vbNewLine & "Проверьте файл-отчеты в папке назначения и повторите расчет IV.")
                        Else
                            varDateArrayFromReportFileNames = varGetFilterDateArrayFromReportFileNames(varDateArrayFromReportFileNames) 'Фильтруем повторяющиеся даты.
                            varFuturesPriceArray = varGetFuturesPriceArray(varProductCode, varDateArrayFromReportFileNames) 'Создаем массив цен фьючерса.
                            '---------------------------------------------------
                            'Combo-начало.
                            '---------------------------------------------------
                            If varCombo = 1 Then 'Если в комбинации одна серия опционов.
                                ReDim varFrontExpDateArray(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAggIVCall(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAggIVPut(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                'Получаем один массив ближайших дат экспирации.
                                byteTrigger = ClosestExpDatesNoCombo(varExpDateArray, _
                                                                     varDateArrayFromReportFileNames, _
                                                                     varFrontExpDateArray)
                                For i = 1 To UBound(varDateArrayFromReportFileNames)
                                    '-------------------------------------------
                                    'Front-начало.
                                    '-------------------------------------------
                                    byteTrigger = StrikeAndIVArrays(varDirectory, _
                                                                    varReportFileNameArray_pr_options_delta_ru(i), _
                                                                    varProductCode, _
                                                                    varFrontExpDateArray(i), _
                                                                    varCallStrikeArray, _
                                                                    varPutStrikeArray, _
                                                                    varCallIVArray, _
                                                                    varPutIVArray, _
                                                                    boolIsThereExpDate_pr_options_delta_ru) 'Получаем Call и Put массивы страйков и волатильностей.
                                    byteTrigger = VolArrays(varDirectory, _
                                                            varReportFileNameArray_rts_results_ru(i), _
                                                            varProductCode, _
                                                            varFrontExpDateArray(i), _
                                                            varCallVolArray, _
                                                            varPutVolArray, _
                                                            boolIsThereExpDate_rts_results_ru) 'Получаем Call и Put массивы объемов.
                                    'If boolIsThereExpDate_pr_options_delta_ru = False And _
                                    '    boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчетах " & varReportFileNameArray_pr_options_delta_ru(i) & ", " _
                                    '    & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_pr_options_delta_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_pr_options_delta_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'End If
                                    If boolIsThereExpDate_pr_options_delta_ru = False Or _
                                       boolIsThereExpDate_rts_results_ru = False Then 'Если даты экспирации в содержимом файл-отчета нет.
                                        varFrontAggIVCall(i) = 0
                                        varFrontAggIVPut(i) = 0
                                        varFrontAvgAggIV(i) = 0
                                    Else
                                        byteTrigger = BubbleSort2(varCallStrikeArray, varCallIVArray, varCallVolArray) 'Сортируем.
                                        byteTrigger = BubbleSort2(varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varFrontAggIVCall(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        varFrontAggIVPut(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varFrontAvgAggIV(i) = (varFrontAggIVCall(i) + varFrontAggIVPut(i)) / 2
                                    End If
                                    '-------------------------------------------
                                    'Front-конец.
                                    '-------------------------------------------
                                Next i
                                'Выводим.
                                byteTrigger = OutputDateArrayFromReportFileNames(varDateArrayFromReportFileNames, intRow_TradeDate, intColumn_TradeDate)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAggIVCall, intRow_FrontAggIVCall, intColumn_FrontAggIVCall)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAggIVPut, intRow_FrontAggIVPut, intColumn_FrontAggIVPut)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAvgAggIV, intRow_FrontAvgAggIV, intColumn_FrontAvgAggIV)
                            End If
                            If varCombo = 2 Then 'Если в комбинации две серии опционов.
                                ReDim varFrontExpDateArray(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextExpDateArray(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAggIVCall(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAggIVPut(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextAggIVCall(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextAggIVPut(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                ReDim varComboAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                'Получаем два массива ближайших дат экспирации.
                                byteTrigger = ClosestExpDatesCombo2(varExpDateArray, _
                                                                    varDateArrayFromReportFileNames, _
                                                                    varFrontExpDateArray, _
                                                                    varNextExpDateArray)
                                For i = 1 To UBound(varDateArrayFromReportFileNames)
                                    '-------------------------------------------
                                    'Front-начало.
                                    '-------------------------------------------
                                    byteTrigger = StrikeAndIVArrays(varDirectory, _
                                                                    varReportFileNameArray_pr_options_delta_ru(i), _
                                                                    varProductCode, _
                                                                    varFrontExpDateArray(i), _
                                                                    varCallStrikeArray, _
                                                                    varPutStrikeArray, _
                                                                    varCallIVArray, _
                                                                    varPutIVArray, _
                                                                    boolIsThereExpDate_pr_options_delta_ru)
                                    byteTrigger = VolArrays(varDirectory, _
                                                            varReportFileNameArray_rts_results_ru(i), _
                                                            varProductCode, _
                                                            varFrontExpDateArray(i), _
                                                            varCallVolArray, _
                                                            varPutVolArray, _
                                                            boolIsThereExpDate_rts_results_ru)
                                    'If boolIsThereExpDate_pr_options_delta_ru = False And _
                                    '    boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчетах " & varReportFileNameArray_pr_options_delta_ru(i) & ", " _
                                    '    & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_pr_options_delta_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_pr_options_delta_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'End If
                                    If boolIsThereExpDate_pr_options_delta_ru = False Or _
                                       boolIsThereExpDate_rts_results_ru = False Then
                                        varFrontAggIVCall(i) = 0
                                        varFrontAggIVPut(i) = 0
                                        varFrontAvgAggIV(i) = 0
                                    Else
                                        byteTrigger = BubbleSort2(varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        byteTrigger = BubbleSort2(varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varFrontAggIVCall(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        varFrontAggIVPut(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varFrontAvgAggIV(i) = (varFrontAggIVCall(i) + varFrontAggIVPut(i)) / 2
                                    End If
                                    '-------------------------------------------
                                    'Front-конец.
                                    '-------------------------------------------
                                    '-------------------------------------------
                                    'Next-начало.
                                    '-------------------------------------------
                                    byteTrigger = StrikeAndIVArrays(varDirectory, _
                                                                    varReportFileNameArray_pr_options_delta_ru(i), _
                                                                    varProductCode, _
                                                                    varNextExpDateArray(i), _
                                                                    varCallStrikeArray, _
                                                                    varPutStrikeArray, _
                                                                    varCallIVArray, _
                                                                    varPutIVArray, _
                                                                    boolIsThereExpDate_pr_options_delta_ru)
                                    byteTrigger = VolArrays(varDirectory, _
                                                            varReportFileNameArray_rts_results_ru(i), _
                                                            varProductCode, _
                                                            varNextExpDateArray(i), _
                                                            varCallVolArray, _
                                                            varPutVolArray, _
                                                            boolIsThereExpDate_rts_results_ru)
                                    'If boolIsThereExpDate_pr_options_delta_ru = False And _
                                    '    boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varNextExpDateArray(i) & _
                                    '    " отсутствует в отчетах " & varReportFileNameArray_pr_options_delta_ru(i) & ", " _
                                    '    & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_pr_options_delta_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varNextExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_pr_options_delta_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varNextExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'End If
                                    If boolIsThereExpDate_pr_options_delta_ru = False Or _
                                       boolIsThereExpDate_rts_results_ru = False Then
                                        varNextAggIVCall(i) = 0
                                        varNextAggIVPut(i) = 0
                                        varNextAvgAggIV(i) = 0
                                    Else
                                        byteTrigger = BubbleSort2(varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        byteTrigger = BubbleSort2(varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varNextAggIVCall(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        varNextAggIVPut(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varNextAvgAggIV(i) = (varNextAggIVCall(i) + varNextAggIVPut(i)) / 2
                                    End If
                                    '-------------------------------------------
                                    'Next-конец.
                                    '-------------------------------------------
                                    varComboAvgAggIV(i) = (varFrontAvgAggIV(i) + varNextAvgAggIV(i)) / 2
                                Next i
                                byteTrigger = OutputDateArrayFromReportFileNames(varDateArrayFromReportFileNames, intRow_TradeDate, intColumn_TradeDate)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAggIVCall, intRow_FrontAggIVCall, intColumn_FrontAggIVCall)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAggIVPut, intRow_FrontAggIVPut, intColumn_FrontAggIVPut)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAvgAggIV, intRow_FrontAvgAggIV, intColumn_FrontAvgAggIV)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varNextAggIVCall, intRow_NextAggIVCall, intColumn_NextAggIVCall)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varNextAggIVPut, intRow_NextAggIVPut, intColumn_NextAggIVPut)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varNextAvgAggIV, intRow_NextAvgAggIV, intColumn_NextAvgAggIV)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varComboAvgAggIV, intRow_ComboAvgAggIV, intColumn_ComboAvgAggIV)
                            End If
                            If varCombo = 3 Then 'Если в комбинации три серии опционов.
                                ReDim varFrontExpDateArray(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextExpDateArray(UBound(varDateArrayFromReportFileNames))
                                ReDim varBackExpDateArray(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAggIVCall(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAggIVPut(UBound(varDateArrayFromReportFileNames))
                                ReDim varFrontAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextAggIVCall(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextAggIVPut(UBound(varDateArrayFromReportFileNames))
                                ReDim varNextAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                ReDim varBackAggIVCall(UBound(varDateArrayFromReportFileNames))
                                ReDim varBackAggIVPut(UBound(varDateArrayFromReportFileNames))
                                ReDim varBackAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                ReDim varComboAvgAggIV(UBound(varDateArrayFromReportFileNames))
                                'Получаем три массива ближайших дат экспирации.
                                byteTrigger = ClosestExpDatesCombo3(varExpDateArray, _
                                                                    varDateArrayFromReportFileNames, _
                                                                    varFrontExpDateArray, _
                                                                    varNextExpDateArray, _
                                                                    varBackExpDateArray)
                                For i = 1 To UBound(varDateArrayFromReportFileNames)
                                    '-------------------------------------------
                                    'Front-начало.
                                    '-------------------------------------------
                                    byteTrigger = StrikeAndIVArrays(varDirectory, _
                                                                    varReportFileNameArray_pr_options_delta_ru(i), _
                                                                    varProductCode, _
                                                                    varFrontExpDateArray(i), _
                                                                    varCallStrikeArray, _
                                                                    varPutStrikeArray, _
                                                                    varCallIVArray, _
                                                                    varPutIVArray, _
                                                                    boolIsThereExpDate_pr_options_delta_ru)
                                    byteTrigger = VolArrays(varDirectory, _
                                                            varReportFileNameArray_rts_results_ru(i), _
                                                            varProductCode, _
                                                            varFrontExpDateArray(i), _
                                                            varCallVolArray, _
                                                            varPutVolArray, _
                                                            boolIsThereExpDate_rts_results_ru)
                                    'If boolIsThereExpDate_pr_options_delta_ru = False And _
                                    '    boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчетах " & varReportFileNameArray_pr_options_delta_ru(i) & ", " _
                                    '    & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_pr_options_delta_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_pr_options_delta_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varFrontExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'End If
                                    If boolIsThereExpDate_pr_options_delta_ru = False Or _
                                       boolIsThereExpDate_rts_results_ru = False Then
                                        varFrontAggIVCall(i) = 0
                                        varFrontAggIVPut(i) = 0
                                        varFrontAvgAggIV(i) = 0
                                    Else
                                        byteTrigger = BubbleSort2(varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        byteTrigger = BubbleSort2(varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varFrontAggIVCall(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        varFrontAggIVPut(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varFrontAvgAggIV(i) = (varFrontAggIVCall(i) + varFrontAggIVPut(i)) / 2
                                    End If
                                    '-------------------------------------------
                                    'Front-конец.
                                    '-------------------------------------------
                                    '-------------------------------------------
                                    'Next-начало.
                                    '-------------------------------------------
                                    byteTrigger = StrikeAndIVArrays(varDirectory, _
                                                                    varReportFileNameArray_pr_options_delta_ru(i), _
                                                                    varProductCode, _
                                                                    varNextExpDateArray(i), _
                                                                    varCallStrikeArray, _
                                                                    varPutStrikeArray, _
                                                                    varCallIVArray, _
                                                                    varPutIVArray, _
                                                                    boolIsThereExpDate_pr_options_delta_ru)
                                    byteTrigger = VolArrays(varDirectory, _
                                                            varReportFileNameArray_rts_results_ru(i), _
                                                            varProductCode, _
                                                            varNextExpDateArray(i), _
                                                            varCallVolArray, _
                                                            varPutVolArray, _
                                                            boolIsThereExpDate_rts_results_ru)
                                    'If boolIsThereExpDate_pr_options_delta_ru = False And _
                                    '    boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varNextExpDateArray(i) & _
                                    '    " отсутствует в отчетах " & varReportFileNameArray_pr_options_delta_ru(i) & ", " _
                                    '    & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_pr_options_delta_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varNextExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_pr_options_delta_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varNextExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'End If
                                    If boolIsThereExpDate_pr_options_delta_ru = False Or _
                                       boolIsThereExpDate_rts_results_ru = False Then
                                        varNextAggIVCall(i) = 0
                                        varNextAggIVPut(i) = 0
                                        varNextAvgAggIV(i) = 0
                                    Else
                                        byteTrigger = BubbleSort2(varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        byteTrigger = BubbleSort2(varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varNextAggIVCall(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        varNextAggIVPut(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varNextAvgAggIV(i) = (varNextAggIVCall(i) + varNextAggIVPut(i)) / 2
                                    End If
                                    '-------------------------------------------
                                    'Next-конец.
                                    '-------------------------------------------
                                    '-------------------------------------------
                                    'Back-начало.
                                    '-------------------------------------------
                                    byteTrigger = StrikeAndIVArrays(varDirectory, _
                                                                    varReportFileNameArray_pr_options_delta_ru(i), _
                                                                    varProductCode, _
                                                                    varBackExpDateArray(i), _
                                                                    varCallStrikeArray, _
                                                                    varPutStrikeArray, _
                                                                    varCallIVArray, _
                                                                    varPutIVArray, _
                                                                    boolIsThereExpDate_pr_options_delta_ru)
                                    byteTrigger = VolArrays(varDirectory, _
                                                            varReportFileNameArray_rts_results_ru(i), _
                                                            varProductCode, _
                                                            varBackExpDateArray(i), _
                                                            varCallVolArray, _
                                                            varPutVolArray, _
                                                            boolIsThereExpDate_rts_results_ru)
                                    'If boolIsThereExpDate_pr_options_delta_ru = False And _
                                    '   boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varBackExpDateArray(i) & _
                                    '    " отсутствует в отчетах " & varReportFileNameArray_pr_options_delta_ru(i) & ", " _
                                    '    & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_pr_options_delta_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varBackExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_pr_options_delta_ru(i) & ".")
                                    'ElseIf boolIsThereExpDate_rts_results_ru = False Then
                                    '    MsgBox ("Оба-на!!!" _
                                    '    & vbNewLine & "Дата экспирации " & varBackExpDateArray(i) & _
                                    '    " отсутствует в отчете " & varReportFileNameArray_rts_results_ru(i) & ".")
                                    'End If
                                    If boolIsThereExpDate_pr_options_delta_ru = False Or _
                                       boolIsThereExpDate_rts_results_ru = False Then
                                        varBackAggIVCall(i) = 0
                                        varBackAggIVPut(i) = 0
                                        varBackAvgAggIV(i) = 0
                                    Else
                                        byteTrigger = BubbleSort2(varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        byteTrigger = BubbleSort2(varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varBackAggIVCall(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varCallStrikeArray, varCallIVArray, varCallVolArray)
                                        varBackAggIVPut(i) = varAggIV(varFuturesPriceArray(i), varPercentage, varPutStrikeArray, varPutIVArray, varPutVolArray)
                                        varBackAvgAggIV(i) = (varBackAggIVCall(i) + varBackAggIVPut(i)) / 2
                                    End If
                                    '-------------------------------------------
                                    'Back-конец.
                                    '-------------------------------------------
                                    varComboAvgAggIV(i) = (varFrontAvgAggIV(i) + varNextAvgAggIV(i) + varBackAvgAggIV(i)) / 3
                                Next i
                                byteTrigger = OutputDateArrayFromReportFileNames(varDateArrayFromReportFileNames, intRow_TradeDate, intColumn_TradeDate)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAggIVCall, intRow_FrontAggIVCall, intColumn_FrontAggIVCall)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAggIVPut, intRow_FrontAggIVPut, intColumn_FrontAggIVPut)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varFrontAvgAggIV, intRow_FrontAvgAggIV, intColumn_FrontAvgAggIV)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varNextAggIVCall, intRow_NextAggIVCall, intColumn_NextAggIVCall)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varNextAggIVPut, intRow_NextAggIVPut, intColumn_NextAggIVPut)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varNextAvgAggIV, intRow_NextAvgAggIV, intColumn_NextAvgAggIV)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varBackAggIVCall, intRow_BackAggIVCall, intColumn_BackAggIVCall)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varBackAggIVPut, intRow_BackAggIVPut, intColumn_BackAggIVPut)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varBackAvgAggIV, intRow_BackAvgAggIV, intColumn_BackAvgAggIV)
                                byteTrigger = OutputAggIV(varDateArrayFromReportFileNames, varComboAvgAggIV, intRow_ComboAvgAggIV, intColumn_ComboAvgAggIV)
                                '-----------------------------------------------
                                'Combo-конец.
                                '-----------------------------------------------
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub
'-------------------------------------------------------------------------------
'Схема исполнения.
'-------------------------------------------------------------------------------
'->varExpDateArray - массив дат экспираций опционов.
'->varDateArrayFromReportFileNames - массив расчетных дат.
'->varReportFileNameArray_pr_options_delta_ru - массив имен файл-отчетов pr-options-delta-ru.
'->varReportFileNameArray_rts_results_ru - массив имен файл-отчетов rts-results-ru.
'->varFuturesPriceArray - массив цен фьючерса.
'->Combo=1 - если в комбинации одна серия опционов.
'  ->varFrontExpDateArray
'  ->varFrontAggIVCall
'  ->varFrontAggIVPut
'  ->varFrontAvgAggIV
'->Combo=2 - если в комбинации две серии опционов.
'  ->varFrontExpDateArray
'  ->varNextExpDateArray
'  ->varFrontAggIVCall
'  ->varFrontAggIVPut
'  ->varFrontAvgAggIV
'  ->varNextAggIVCall
'  ->varNextAggIVPut
'  ->varNextAvgAggIV
'  ->varComboAvgAggIV
'->Combo=3 - если в комбинации три серии опционов.
'  ->varFrontExpDateArray
'  ->varNextExpDateArray
'  ->varBackExpDateArray
'  ->varFrontAggIVCall
'  ->varFrontAggIVPut
'  ->varFrontAvgAggIV
'  ->varNextAggIVCall
'  ->varNextAggIVPut
'  ->varNextAvgAggIV
'  ->varBackAggIVCall
'  ->varBackAggIVPut
'  ->varBackAvgAggIV
'  ->varComboAvgAggIV
'->Вывод на экран.

'-------------------------------------------------------------------------------
'Буфер запасных функций.
'-------------------------------------------------------------------------------
''Функция находит даты в именах файлов и создает массив дат.
'Private Function varCreateDateArrayFromReports(varFileNameArray As Variant) As Variant
'    On Error GoTo Errors
'    Dim strItem1 As String, strItem2 As String
'    Dim intQtyItems As Integer, intCounter As Integer, i As Integer, j As Integer
'    Dim varStartingPointForDate As Variant, varEndPointForDate As Variant
'    Dim dateDateFromReport As Date, dateDateArrayFromReports() As Date
'    ReDim dateDateArrayFromReports((UBound(varFileNameArray)) / 2) 'Делим на 2, т.к.
'    'на 2 разных файла (pr-options-delta-ru и rts-results-ru) приходится две одинаковые даты.
'    intCounter = 1
'    'Начинаем просмотр всех имен.
'    For i = 1 To UBound(varFileNameArray)
'        'Определяем количество элементов в имени.
'        intQtyItems = Len(varFileNameArray(i))
'        'Начинаем поиск даты в имени.
'        For j = 1 To intQtyItems
'            strItem1 = Mid(varFileNameArray(i), j, 1) 'Берем элемент.
'            If Asc(strItem1) >= 48 And Asc(strItem1) <= 57 Then 'Если элемент является числом (Asc(0)=48...Asc(9)=57).
'                If IsEmpty(varStartingPointForDate) = True Then 'Если метка начала даты пуста.
'                    varStartingPointForDate = j
'                ElseIf j <> intQtyItems Then 'Если элемент не последний.
'                    strItem2 = Mid(varFileNameArray(i), j + 1, 1) 'Берем соседний элемент справа.
'                    If Asc(strItem2) < 48 And Asc(strItem2) <> Asc(".") Or Asc(strItem2) > 57 Then 'Если элемент не число и не точка.
'                        varEndPointForDate = j
'                        Exit For
'                    End If
'                Else: varEndPointForDate = j 'Если элемент последний (здесь мы знаем, что элемент число).
'                End If
'            End If
'        Next j
'        dateDateFromReport = Mid(varFileNameArray(i), varStartingPointForDate, (varEndPointForDate + 1) - varStartingPointForDate) 'Получаем дату.
'        If IsDate(dateDateFromReport) = False Then 'Проверяем дату. Если это не дата.
'            MsgBox ("Оба-на!!!" _
'            & vbNewLine & "Где-то некорректно указана дата." _
'            & vbNewLine & "Проверьте файлы в папке назначения и повторите расчет IV.")
'            Exit Function
'        End If
'        If i = 1 Then 'Если имя первое.
'            dateDateArrayFromReports(1) = dateDateFromReport
'        ElseIf dateDateFromReport <> dateDateArrayFromReports(intCounter) Then 'Если текущая дата не равна предыдущей.
'            intCounter = intCounter + 1
'            dateDateArrayFromReports(intCounter) = dateDateFromReport
'            GoTo Ends: 'Если на 2 разных файла (pr-options-delta-ru и rts-results-ru) приходится две разные даты.
'Errors:
'            MsgBox ("Оба-на!!!" _
'            & vbNewLine & "Где-то некорректно указана дата." _
'            & vbNewLine & "Проверьте файлы в папке назначения и повторите расчет IV.")
'            Exit Function
'Ends:
'        End If
'    Next i
'    varCreateDateArrayFromReports = dateDateArrayFromReports
'End Function

''Функция сортирует массив дат по возрастанию (метод Пузырек).
'Private Function varSortBM(varDateArrayFromReportFileNames As Variant) As Variant
'    Dim i As Integer
'    Dim boolSwitch As Boolean
'    Dim varTmp As Variant
'    Do
'        boolSwitch = 0
'        For i = 1 To UBound(varDateArrayFromReportFileNames) - 1
'            If DateValue(varDateArrayFromReportFileNames(i + 1)) < DateValue(varDateArrayFromReportFileNames(i)) Then
'                boolSwitch = 1
'                varTmp = varDateArrayFromReportFileNames(i)
'                varDateArrayFromReportFileNames(i) = varDateArrayFromReportFileNames(i + 1)
'                varDateArrayFromReportFileNames(i + 1) = varTmp
'            End If
'        Next i
'        If boolSwitch = 0 Then Exit Do
'    Loop
'    varSortBM = varDateArrayFromReportFileNames
'End Function
