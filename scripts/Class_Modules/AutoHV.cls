VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "HCI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Option Compare Text
Dim byteTrigger As Byte

'Функция тестирует попадания.
Private Function boolHitTest(varArrow As Variant, _
                             varTarget As Variant) As Boolean
    Dim i As Integer, _
        lngUbound As Long
    If IsObject(varTarget) = True Then
        lngUbound = varTarget.Count
    Else lngUbound = UBound(varTarget)
        For i = 1 To lngUbound
            If varArrow = varTarget(i) Then
                boolHitTest = True
                Exit Function
            End If
        Next
End Function

'Функция дает адрес ячейки входного параметра.
Private Function GetAdressParam(intRow As Integer, _
                                intColumn As Integer, _
                                strParamName As String)
    Dim i As Integer, j As Integer
    For i = 1 To 100
        For j = 1 To 100
            If IsError(Worksheets("AutoHV").Cells(i, j).Value) = False Then
                If Worksheets("AutoHV").Cells(i, j).Value = strParamName Then
                    intRow = i
                    intColumn = j
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

'Функция дает количество файл-отчетов в папке назначения.
Private Function intGetQtyReportFiles(varDirectory As Variant) As Integer
    Dim strReportFileName As String, _
        intCounter As Integer
    'Получение имени первого файл-отчета (образец первого вызова).
    strReportFileName = Dir(varDirectory, vbReadOnly + vbHidden + vbSystem)
    If strReportFileName = "" Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "В папке назначения файл-отчетов нет." _
        & vbNewLine & "Поместите файл-отчеты в папку и повторите расчет HV.")
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
    Dim strReportFileNameArray() As String, _
        intCounter As Integer
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

'Функция ловит короткий код фьючерсного контракта, используя регулярное выражение.
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
Private Function varGetShortCodeFut(varReportFileName As Variant) As Variant
    'Dim Re As Object
    'Set Re = CreateObject("vbscript.regexp")
    Dim Re As RegExp
    Set Re = New RegExp
    Re.Pattern = "[A-Z]+[A-Za-z]+[FGHJKMNQUVXZ]+[0-9]" '2 символа кода базового актива + символ месяца + год.
    If Re.test(varReportFileName) Then varGetShortCodeFut = Re.Execute(varReportFileName)(0)
    Set Re = Nothing
End Function

'Функция дает массив коротких кодов.
Private Function varGetShortCodeFutArray(varReportFileNameArray As Variant) As Variant
    Dim i As Integer, _
        varShortCodeFutArray() As Variant
    ReDim varShortCodeFutArray(UBound(varReportFileNameArray))
    For i = 1 To UBound(varReportFileNameArray)
        varShortCodeFutArray(i) = varGetShortCodeFut(varReportFileNameArray(i))
    Next i
    varGetShortCodeFutArray = varShortCodeFutArray
End Function

'From To functions <<<
'Sets lower bound for array subscript to 1
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

'Returns a Collection which exactly matches the given Array
' Note: This function is not safe for concurrent modification.
Private Function colFromArray(varA() As Variant) As Collection
    Dim varElement As Variant, _
        col As Collection
    Set col = New Collection
    For Each varElement In varA
        col.Add varElement
    Next varElement
    Set colFromArray = col
End Function
'From To functions >>>

'Sort functions <<<
'Функция сортирует массив по возрастанию (метод Пузырек).
Private Function BubbleSort(varArray As Variant)
    Dim i As Integer, _
        boolSwitch As Boolean, _
        varTmp As Variant
    Do
        boolSwitch = 0
        For i = 1 To UBound(varArray) - 1
            If varArray(i + 1) < varArray(i) Then
                boolSwitch = 1
                varTmp = varArray(i)
                varArray(i) = varArray(i + 1)
                varArray(i + 1) = varTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
End Function

Private Function BubbleSort2(varArray As Variant)
    Dim i As Integer, _
        boolSwitch As Boolean, _
        varTmp As Variant
    Do
        boolSwitch = 0
        For i = 1 To UBound(varArray) - 1
            If Mid(varArray(i + 1), 3, 1) < Mid(varArray(i), 3, 1) Then 'Сортировка по месяцу исполнения.
                boolSwitch = 1
                varTmp = varArray(i)
                varArray(i) = varArray(i + 1)
                varArray(i + 1) = varTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
End Function

Private Function BubbleSort3(col As Collection)
    Dim varItem As Variant, _
        i As Long, j As Long, _
        varTemp As Variant
    'Two loops to bubble sort
    For i = 1 To col.Count - 1
        For j = i + 1 To col.Count
            If col(i) > col(j) Then
                'store the lesser item
                varTemp = col(j)
                'remove the lesser item
                col.Remove j
                're-add the lesser item before the
                'greater Item
                col.Add varTemp, varTemp, i
            End If
        Next j
    Next i
End Function
'Sort functions >>>

'Функция дает коллекцию лет исполнения.
Private Function colGetFutExecYearArray(varShortCodeFutArray As Variant) As Collection
    Dim i As Integer, _
        colScan As New Collection
    For i = 1 To UBound(varShortCodeFutArray)
        byteTrigger = boolHitTest(Mid(varShortCodeFutArray(i), 4, 1), colScan) 'Сравниваем год исполнения с уже просмотренными.
        If byteTrigger = False Then ' Если совпадения нет.
            colScan.Add Mid(varShortCodeFutArray(i), 4, 1) 'Добавляем в коллекцию.
        End If
    Next i
    Set colGetFutExecYearArray = colScan
End Function

'Функция сортирует массив коротких кодов.
Private Function varSortShortCodeFutArray(varShortCodeFutArray As Variant, _
                                          colFutExecYearArray As Collection) As Variant
    Dim i As Integer, j As Integer, _
        varTmpArray As Variant, _
        colSortShortCodeFutArray As New Collection
    For i = 1 To colFutExecYearArray.Count
        varTmpArray = Filter(varShortCodeFutArray, colFutExecYearArray(i)) 'Фильтруем группу коротких кодов.
        varTmpArray = varToOptionBase1(varTmpArray)
        byteTrigger = BubbleSort2(varTmpArray) 'Сортируем по месяцу исполнения.
        For j = 1 To UBound(varTmpArray)
            colSortShortCodeFutArray.Add varTmpArray(j) 'Добавляем в коллекцию.
        Next j
    Next i
    varShortCodeFutArray = varToArray(colSortShortCodeFutArray) 'Получаем сортированный массив коротких кодов.
    varSortShortCodeFutArray = varShortCodeFutArray
End Function

'Функция сортирует массив имен файл-отчетов.
Private Function varSortReportFileNameArray(varReportFileNameArray As Variant, _
                                            varShortCodeFutArray As Variant) As Variant
    Dim i As Integer, _
        strReportFileName() As String, _
        colSortReportFileNameArray As New Collection
    For i = 1 To UBound(varReportFileNameArray)
        strReportFileName = Filter(varReportFileNameArray, varShortCodeFutArray(i)) 'Фильтруем имя файл-отчета.
        colSortReportFileNameArray.Add strReportFileName(0) 'Добавляем в коллекцию.
    Next i
    varReportFileNameArray = varToArray(colSortReportFileNameArray) 'Получаем сортированный массив имен файл-отчетов.
    varSortReportFileNameArray = varReportFileNameArray
End Function

'Главная функция сортировки массива имен файл-отчетов (групповая сортировка).
Private Function varMainSortReportFileNameArray(varReportFileNameArray As Variant) As Variant
    Dim i As Integer, _
        varShortCodeFutArray As Variant, _
        colFutExecYearArray As New Collection
    varShortCodeFutArray = varGetShortCodeFutArray(varReportFileNameArray) 'Получаем массив коротких кодов.
    Set colFutExecYearArray = colGetFutExecYearArray(varShortCodeFutArray) 'Получаем коллекцию лет исполнения.
    byteTrigger = BubbleSort3(colFutExecYearArray) 'Сортируем коллекцию лет исполнения.
    varShortCodeFutArray = varSortShortCodeFutArray(varShortCodeFutArray, colFutExecYearArray) 'Получаем сортированный массив коротких кодов.
    varReportFileNameArray = varSortReportFileNameArray(varReportFileNameArray, varShortCodeFutArray) 'Получаем сортированный массив имен файл-отчетов.
    varMainSortReportFileNameArray = varReportFileNameArray
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

'Функция возвращает 0, если символ пустой строки "", иначе само значение файл-отчета.
Private Function dblReportFileValue(varContentFilterFromReportFile) As Double
    If varContentFilterFromReportFile = "" Then
        dblReportFileValue = CDbl(0)
    Else dblReportFileValue = CDbl(varContentFilterFromReportFile)
End Function

'Функция дает двумерный массив MarketData (1 to x,1 to 6). 1 to x - кол-во строк,
'1 to 6 - количество столбцов (TradeDate, Settle, Open, High, Low, Close).
Private Function varGetMarketDataArray(varDirectory As Variant, _
                                       varReportFileName As Variant) As Variant
    Dim i As Integer, _
        varContentFromReportFile As Variant, _
        varContentFilterFromReportFile As Variant, _
        varMarketDataArray() As Variant
    varContentFromReportFile = strGetContentFromReportFile(varDirectory, varReportFileName) 'Считываем содержимое файл-отчета в строку.
    varContentFromReportFile = Split(varContentFromReportFile, vbLf) 'Разбиваем содержимое файл-отчета на строки. Разделитель - символ LF - Chr(10) - vbLf.
    varContentFromReportFile = varToOptionBase1(varContentFromReportFile) 'Приводим массив в соответствие с оператором Option Base 1.
    ReDim varMarketDataArray(1 To UBound(varContentFromReportFile) - 2, 1 To 6) 'Вычитаем 1 строку "шапка" и последнюю с символом пустой строки "".
    For i = 2 To UBound(varContentFromReportFile) - 1
        varContentFilterFromReportFile = Split(varContentFromReportFile(i), ";") 'Т.к. функция Split определяет нижнюю границу массива=0, то поля смещаются на 1.
        varMarketDataArray(i - 1, 1) = varContentFilterFromReportFile(0) 'TradeDate
        varMarketDataArray(i - 1, 2) = dblReportFileValue(varContentFilterFromReportFile(2)) 'Settle.
        varMarketDataArray(i - 1, 3) = dblReportFileValue(varContentFilterFromReportFile(3)) 'Open.
        varMarketDataArray(i - 1, 4) = dblReportFileValue(varContentFilterFromReportFile(4)) 'High.
        varMarketDataArray(i - 1, 5) = dblReportFileValue(varContentFilterFromReportFile(5)) 'Low.
        varMarketDataArray(i - 1, 6) = dblReportFileValue(varContentFilterFromReportFile(6)) 'Close.
    Next i
    varGetMarketDataArray = varMarketDataArray
End Function

'Функция дает массив массивов MarketData (x)(1 to y, 1 to 6), x - кол-во
'файл-отчетов, 1 to y - кол-во строк, 1 to 6 кол-во столбцов.
Private Function varGetMarketDataJaggArray(varDirectory As Variant, _
                                           varReportFileNameArray As Variant) As Variant
    Dim i As Integer, _
        varMarketDataArray As Variant, _
        varMarketDataJaggArray() As Variant
    ReDim varMarketDataJaggArray(UBound(varReportFileNameArray))
    For i = 1 To UBound(varReportFileNameArray)
        varMarketDataArray = varGetMarketDataArray(varDirectory, varReportFileNameArray(i)) 'Получаем MarketData из файл-отчета.
        varMarketDataJaggArray(i) = varMarketDataArray 'Добавляем массив в массив.
    Next i
    varGetMarketDataJaggArray = varMarketDataJaggArray
End Function

'Схема соединения ([(e2+1)to(b1-1)]+[(b1toe1)+(b2toe2)]+[(e2+1)to(b1-1)]+
'[(b1toe1)+(b2toe2)]+...+[(e2+1)to(b1-1)]).
'Точки соединения первого порядка находятся от первого массива MarketData
'до предпоследнего (последняя точка соот. предпоследнему массиву).
'Точки соединения вторго порядка находятся от второго массива MarketData
'до последнего (последняя точка соот. последнему массиву).
'Начальная нулевая точка конца соединения второго порядка=0 т.е. 1 (0+1).
'Конечная нулевая точка начала соединения первого порядка=номеру последней
'строки+1 т.е. номеру последней строки (номер последней строки-1).
'   A_________0e2
'   B
'b1_C A________b2
'e1_D B________e2
'  b1_C A______b2
'  e1_D B______e2
'    b1_С A____b2
'    e1_D B____e2
'         С
'0b1______D
'Функция дает массив номеров строк конца соединения (точки соединения первого порядка).
Private Function varGetEndJoin1stRowNumArray(varMarketDataJaggArray As Variant) As Variant
    Dim i As Integer, _
        intEndJoin1stRowNumArray() As Integer
    ReDim intEndJoin1stRowNumArray(UBound(varMarketDataJaggArray) - 1)
    For i = 1 To UBound(varMarketDataJaggArray) - 1 'От первого до предпоследнего.
        intEndJoin1stRowNumArray(i) = UBound(varMarketDataJaggArray(i), 1)
    Next i
    varGetEndJoin1stRowNumArray = intEndJoin1stRowNumArray
End Function

'Функция дает массив номеров строк начала соединения (точки соединения).
Private Function varGetBegJoinRowNumArray(varEndJoinRowNumArray As Variant, _
                                          byteJoinDepth As Byte) As Variant
    Dim i As Integer, _
        intBegJoinRowNumArray() As Integer
    ReDim intBegJoinRowNumArray(UBound(varEndJoinRowNumArray)) 'Кол-во строк начала соединения соот. кол-ву строк конца.
    For i = 1 To UBound(varEndJoinRowNumArray)
        intBegJoinRowNumArray(i) = (varEndJoinRowNumArray(i) - byteJoinDepth) + 1
    Next i
    varGetBegJoinRowNumArray = intBegJoinRowNumArray
End Function

'Функция дает массив дат соединения по номерам строк.
Private Function varGetJoinDateArray(varMarketDataJaggArray As Variant, _
                                     varJoinRowNumArray As Variant) As Variant
    Dim i As Integer, _
        strJoinDateArray() As String
    ReDim strJoinDateArray(UBound(varJoinRowNumArray)) 'Кол-во дат соот. кол-ву строк.
    For i = 1 To UBound(varMarketDataJaggArray) - 1 'От первого до предпоследнего.
        strJoinDateArray(i) = varMarketDataJaggArray(i)(varJoinRowNumArray(i), 1) 'Первая дата соот. первому массиву MarketData, последняя - предпоследнему.
    Next i
    varGetJoinDateArray = strJoinDateArray
End Function

'Функция дает массив номеров строк соединения по датам (точки соединения второго порядка).
Private Function varGetJoinRowNumArray(varMarketDataJaggArray As Variant, _
                                       varJoinDateArray As Variant) As Variant
    Dim i As Integer, j As Integer, _
        intJoinRowNumArray() As Integer
    ReDim intJoinRowNumArray(UBound(varJoinDateArray)) 'Кол-во строк соот. кол-ву дат.
    For i = 2 To UBound(varMarketDataJaggArray) 'От второго до последнего.
        For j = 1 To UBound(varMarketDataJaggArray(i), 1)
            If varMarketDataJaggArray(i)(j, 1) = varJoinDateArray(i - 1) Then intJoinRowNumArray(i - 1) = j 'Сравниваем дату следующего массива MarketData с датой предыдущего.
            Next j
    Next i
    varGetJoinRowNumArray = intJoinRowNumArray
End Function

'Функция дает массив массивов номеров строк соединения (все точки соединения).
Private Function varGetJoinRowNumJaggArray(varMarketDataJaggArray As Variant, _
                                           byteJoinDepth As Byte) As Variant
    Dim varEndJoin1stRowNumArray As Variant, _
        varBegJoin1stRowNumArray As Variant, _
        varEndJoinDateArray As Variant, _
        varEndJoin2ndRowNumArray As Variant, _
        varBegJoin2ndRowNumArray As Variant, _
        varJoinRowNumJaggArray(4) As Variant
    varEndJoin1stRowNumArray = varGetEndJoin1stRowNumArray(varMarketDataJaggArray) 'Получаем номера строк конца соединения (первого порядка).
    varBegJoin1stRowNumArray = varGetBegJoinRowNumArray(varEndJoin1stRowNumArray, byteJoinDepth) 'Получаем номера строк начала соединения (первого порядка).
    varEndJoinDateArray = varGetJoinDateArray(varMarketDataJaggArray, varEndJoin1stRowNumArray) 'Получаем даты конца соединения.
    varEndJoin2ndRowNumArray = varGetJoinRowNumArray(varMarketDataJaggArray, varEndJoinDateArray) 'Получаем номера строк конца соединения (второго порядка).
    varBegJoin2ndRowNumArray = varGetBegJoinRowNumArray(varEndJoin2ndRowNumArray, byteJoinDepth) 'Получаем номера строк начала соединения (второго порядка).
    varJoinRowNumJaggArray(1) = varBegJoin1stRowNumArray
    varJoinRowNumJaggArray(2) = varEndJoin1stRowNumArray
    varJoinRowNumJaggArray(3) = varBegJoin2ndRowNumArray
    varJoinRowNumJaggArray(4) = varEndJoin2ndRowNumArray
    varGetJoinRowNumJaggArray = varJoinRowNumJaggArray
End Function

'Функция дает пустой фрагмент соединения.
Private Function varGetEmptyFragment() As Variant
    Dim varEmptyFragment(1, 1 To 6) As Variant
    varEmptyFragment(1, 1) = "31.12.9999"
    varEmptyFragment(1, 2) = Empty
    varEmptyFragment(1, 3) = Empty
    varEmptyFragment(1, 4) = Empty
    varEmptyFragment(1, 5) = Empty
    varEmptyFragment(1, 6) = Empty
    varGetEmptyFragment = varEmptyFragment
End Function

'Функция дает массив осн. фрагмента соединения MarketData.
Private Function varGetBaseFragmentMarketDataArray(varMarketDataArray As Variant, _
                                                   varEndJoin2ndRowNum As Variant, _
                                                   varBegJoin1stRowNum As Variant) As Variant
    Dim i As Integer, _
        varBaseFragmentMarketDataArray() As Variant
    ReDim varBaseFragmentMarketDataArray((varBegJoin1stRowNum - varEndJoin2ndRowNum) - 1, 1 To 6) 'Номер строки начала соединения первого порядка-номер строки конца соединения второго порядка-1.
    For i = varEndJoin2ndRowNum + 1 To varBegJoin1stRowNum - 1 'От номера строки конца соединения второго порядка+1 до номера строки начала соединения первого порядка-1.
        varBaseFragmentMarketDataArray(i - varEndJoin2ndRowNum, 1) = varMarketDataArray(i, 1)
        varBaseFragmentMarketDataArray(i - varEndJoin2ndRowNum, 2) = varMarketDataArray(i, 2)
        varBaseFragmentMarketDataArray(i - varEndJoin2ndRowNum, 3) = varMarketDataArray(i, 3)
        varBaseFragmentMarketDataArray(i - varEndJoin2ndRowNum, 4) = varMarketDataArray(i, 4)
        varBaseFragmentMarketDataArray(i - varEndJoin2ndRowNum, 5) = varMarketDataArray(i, 5)
        varBaseFragmentMarketDataArray(i - varEndJoin2ndRowNum, 6) = varMarketDataArray(i, 6)
    Next i
    varGetBaseFragmentMarketDataArray = varBaseFragmentMarketDataArray
End Function

'Функция дает массив массивов осн. фрагментов соединения MarketData.
Private Function varGetBaseFragmentsMarketDataJaggArray(varMarketDataJaggArray As Variant, _
                                                        varJoinRowNumJaggArray As Variant) As Variant
    Dim i As Integer, _
        varBaseFragmentsMarketDataJaggArray() As Variant
    ReDim varBaseFragmentsMarketDataJaggArray(UBound(varMarketDataJaggArray)) 'Кол-во осн. фрагментов соот. кол-ву массивов MarketData.
    varBaseFragmentsMarketDataJaggArray(1) = varGetBaseFragmentMarketDataArray(varMarketDataJaggArray(1), 0, varJoinRowNumJaggArray(1)(1)) 'Получаем первый осн. фрагмент (см. схему соединения).
    For i = 2 To UBound(varMarketDataJaggArray) - 1 'Получаем осн. фрагменты от второго до предпоследнего (см. схему соединения).
        varBaseFragmentsMarketDataJaggArray(i) = varGetBaseFragmentMarketDataArray(varMarketDataJaggArray(i), varJoinRowNumJaggArray(4)(i - 1), varJoinRowNumJaggArray(1)(i))
    Next i
    'Получаем последний осн. фрагмент (см. схему соединения).
    If varJoinRowNumJaggArray(4)(UBound(varJoinRowNumJaggArray(4))) = UBound(varMarketDataJaggArray(UBound(varMarketDataJaggArray)), 1) Then
        varBaseFragmentsMarketDataJaggArray(UBound(varMarketDataJaggArray)) = varGetEmptyFragment 'Пустой.
    Else
        varBaseFragmentsMarketDataJaggArray(UBound(varMarketDataJaggArray)) = varGetBaseFragmentMarketDataArray(varMarketDataJaggArray(UBound(varMarketDataJaggArray)), _
        varJoinRowNumJaggArray(4)(UBound(varJoinRowNumJaggArray(4))), _
        UBound(varMarketDataJaggArray(UBound(varMarketDataJaggArray)), 1) + 1) 'Непустой.
    End If
    varGetBaseFragmentsMarketDataJaggArray = varBaseFragmentsMarketDataJaggArray
End Function

'Функция дает массив доп. фрагмента соединения MarketData.
Private Function varGetAddFragmentMarketDataArray(varMarketDataArray As Variant, _
                                                  varEndJoinRowNum As Variant, _
                                                  varBegJoinRowNum As Variant) As Variant
    Dim i As Integer, _
        varAddFragmentMarketDataArray() As Variant
    ReDim varAddFragmentMarketDataArray((varEndJoinRowNum - varBegJoinRowNum) + 1, 1 To 6)
    For i = varBegJoinRowNum To varEndJoinRowNum 'От номера строки начала соединения до номера строки конца соединения (начало и конец входят).
        varAddFragmentMarketDataArray(i - (varBegJoinRowNum - 1), 1) = varMarketDataArray(i, 1)
        varAddFragmentMarketDataArray(i - (varBegJoinRowNum - 1), 2) = varMarketDataArray(i, 2)
        varAddFragmentMarketDataArray(i - (varBegJoinRowNum - 1), 3) = varMarketDataArray(i, 3)
        varAddFragmentMarketDataArray(i - (varBegJoinRowNum - 1), 4) = varMarketDataArray(i, 4)
        varAddFragmentMarketDataArray(i - (varBegJoinRowNum - 1), 5) = varMarketDataArray(i, 5)
        varAddFragmentMarketDataArray(i - (varBegJoinRowNum - 1), 6) = varMarketDataArray(i, 6)
    Next i
    varGetAddFragmentMarketDataArray = varAddFragmentMarketDataArray
End Function

'Функция дает массив массивов доп. фрагментов соединения MarketData первого порядка.
Private Function varGetAddFragments1stMarketDataJaggArray(varMarketDataJaggArray As Variant, _
                                                          varJoinRowNumJaggArray As Variant) As Variant
    Dim i As Integer, _
        varAddFragments1stMarketDataJaggArray() As Variant
    ReDim varAddFragments1stMarketDataJaggArray(UBound(varMarketDataJaggArray) - 1)
    For i = 1 To UBound(varMarketDataJaggArray) - 1 'От первого до предпоследнего (см. схему соединения).
        varAddFragments1stMarketDataJaggArray(i) = varGetAddFragmentMarketDataArray(varMarketDataJaggArray(i), varJoinRowNumJaggArray(2)(i), varJoinRowNumJaggArray(1)(i))
    Next i
    varGetAddFragments1stMarketDataJaggArray = varAddFragments1stMarketDataJaggArray
End Function

'Функция дает массив массивов доп. фрагментов соединения MarketData второго порядка.
Private Function varGetAddFragments2ndMarketDataJaggArray(varMarketDataJaggArray As Variant, _
                                                          varJoinRowNumJaggArray As Variant) As Variant
    Dim i As Integer, _
        varAddFragments2ndMarketDataJaggArray() As Variant
    ReDim varAddFragments2ndMarketDataJaggArray(UBound(varMarketDataJaggArray) - 1)
    For i = 2 To UBound(varMarketDataJaggArray) 'От второго до последнего (см. схему соединения).
        varAddFragments2ndMarketDataJaggArray(i - 1) = varGetAddFragmentMarketDataArray(varMarketDataJaggArray(i), varJoinRowNumJaggArray(4)(i - 1), varJoinRowNumJaggArray(3)(i - 1))
    Next i
    varGetAddFragments2ndMarketDataJaggArray = varAddFragments2ndMarketDataJaggArray
End Function

'Функция дает усредненный массив доп. фрагмента соединения MarketData.
Private Function varGetAvgAddFragmentMarketDataArray(varAddFragment1stMarketDataArray As Variant, _
                                                     varAddFragment2ndMarketDataArray As Variant, _
                                                     strJoinMode As String) As Variant
    Dim i As Integer, _
        varAvgAddFragmentMarketDataArray() As Variant, _
        f As WorksheetFunction
    Set f = Application.WorksheetFunction
    ReDim varAvgAddFragmentMarketDataArray(UBound(varAddFragment1stMarketDataArray, 1), 1 To 6)
    For i = 1 To UBound(varAddFragment1stMarketDataArray, 1)
        varAvgAddFragmentMarketDataArray(i, 1) = varAddFragment1stMarketDataArray(i, 1) 'Date
        'If any data point <= 0, GEOMEAN returns the #NUM! error value.
        'Sqr(v1 * v2) is used instead of GeoMean when n = 2.
        varAvgAddFragmentMarketDataArray(i, 2) = IIf(UCase(strJoinMode) = "AM", f.Average(varAddFragment1stMarketDataArray(i, 2), varAddFragment2ndMarketDataArray(i, 2)), _
        Sqr(varAddFragment1stMarketDataArray(i, 2) * varAddFragment2ndMarketDataArray(i, 2)))
        varAvgAddFragmentMarketDataArray(i, 3) = IIf(UCase(strJoinMode) = "AM", f.Average(varAddFragment1stMarketDataArray(i, 3), varAddFragment2ndMarketDataArray(i, 3)), _
        Sqr(varAddFragment1stMarketDataArray(i, 3) * varAddFragment2ndMarketDataArray(i, 3)))
        varAvgAddFragmentMarketDataArray(i, 4) = IIf(UCase(strJoinMode) = "AM", f.Average(varAddFragment1stMarketDataArray(i, 4), varAddFragment2ndMarketDataArray(i, 4)), _
        Sqr(varAddFragment1stMarketDataArray(i, 4) * varAddFragment2ndMarketDataArray(i, 4)))
        varAvgAddFragmentMarketDataArray(i, 5) = IIf(UCase(strJoinMode) = "AM", f.Average(varAddFragment1stMarketDataArray(i, 5), varAddFragment2ndMarketDataArray(i, 5)), _
        Sqr(varAddFragment1stMarketDataArray(i, 5) * varAddFragment2ndMarketDataArray(i, 5)))
        varAvgAddFragmentMarketDataArray(i, 6) = IIf(UCase(strJoinMode) = "AM", f.Average(varAddFragment1stMarketDataArray(i, 6), varAddFragment2ndMarketDataArray(i, 6)), _
        Sqr(varAddFragment1stMarketDataArray(i, 6) * varAddFragment2ndMarketDataArray(i, 6)))
    Next i
    varGetAvgAddFragmentMarketDataArray = varAvgAddFragmentMarketDataArray
End Function

'Функция дает усредненный массив массивов доп. фрагментов соединения MarketData.
Private Function varGetAvgAddFragmentsMarketDataJaggArray(varMarketDataJaggArray As Variant, _
                                                          varJoinRowNumJaggArray As Variant, _
                                                          strJoinMode As String) As Variant
    Dim i As Integer, _
        varAddFragments1stMarketDataJaggArray As Variant, _
        varAddFragments2ndMarketDataJaggArray As Variant, _
        varAvgAddFragmentsMarketDataJaggArray() As Variant
    varAddFragments1stMarketDataJaggArray = varGetAddFragments1stMarketDataJaggArray(varMarketDataJaggArray, varJoinRowNumJaggArray) 'Получаем массив массивов доп. фрагментов соединения MarketData первого порядка.
    varAddFragments2ndMarketDataJaggArray = varGetAddFragments2ndMarketDataJaggArray(varMarketDataJaggArray, varJoinRowNumJaggArray) 'Получаем массив массивов доп. фрагментов соединения MarketData второго порядка.
    ReDim varAvgAddFragmentsMarketDataJaggArray(UBound(varAddFragments1stMarketDataJaggArray))
    For i = 1 To UBound(varAddFragments1stMarketDataJaggArray)
        varAvgAddFragmentsMarketDataJaggArray(i) = varGetAvgAddFragmentMarketDataArray(varAddFragments1stMarketDataJaggArray(i), varAddFragments2ndMarketDataJaggArray(i), strJoinMode)
    Next i
    varGetAvgAddFragmentsMarketDataJaggArray = varAvgAddFragmentsMarketDataJaggArray
End Function

'Функция дает агрегированный размер массива массивов.
Private Function intGetAggSizeJaggArray(varJaggArray As Variant) As Integer
    Dim i As Integer, intAggSizeJaggArray As Integer
    intAggSizeJaggArray = 0
    For i = 1 To UBound(varJaggArray)
        intAggSizeJaggArray = intAggSizeJaggArray + UBound(varJaggArray(i), 1)
    Next i
    intGetAggSizeJaggArray = intAggSizeJaggArray
End Function

'Функция дает номер пустой строки.
Private Function intGetNumEmptyRow(varArray As Variant) As Integer
    Dim i As Integer
    For i = 1 To UBound(varArray, 1)
        If IsEmpty(varArray(i, 1)) Then
            intGetNumEmptyRow = i
            Exit Function
        End If
    Next i
End Function

'Функция выполняет целевое добавление элементов в массив.
Private Function TargetAddition(varBaseArray As Variant, varAddArray As Variant)
    Dim i As Integer, intNumEmptyRow As Integer
    intNumEmptyRow = intGetNumEmptyRow(varBaseArray) 'Получаем начало добавления.
    For i = intNumEmptyRow To intNumEmptyRow + UBound(varAddArray, 1) - 1
        varBaseArray(i, 1) = varAddArray(i - (intNumEmptyRow - 1), 1)
        varBaseArray(i, 2) = varAddArray(i - (intNumEmptyRow - 1), 2)
        varBaseArray(i, 3) = varAddArray(i - (intNumEmptyRow - 1), 3)
        varBaseArray(i, 4) = varAddArray(i - (intNumEmptyRow - 1), 4)
        varBaseArray(i, 5) = varAddArray(i - (intNumEmptyRow - 1), 5)
        varBaseArray(i, 6) = varAddArray(i - (intNumEmptyRow - 1), 6)
    Next i
End Function

'Функция дает соединенный массив MarketData.
Private Function varGetJoinMarketDataArray(strDirectory As String, _
                                           varReportFileNameArray As Variant, _
                                           strJoinMode As String, _
                                           byteJoinDepth As Byte)
    Dim i As Integer, _
        varMarketDataJaggArray As Variant, _
        varJoinRowNumJaggArray As Variant, _
        varBaseFragmentsMarketDataJaggArray As Variant, _
        varAvgAddFragmentsMarketDataJaggArray As Variant, _
        intAggSizeBaseFragmentsMarketDataJaggArray As Integer, _
        intAggSizeAvgAddFragmentsMarketDataJaggArray As Integer, _
        varJoinMarketDataArray() As Variant
    varMarketDataJaggArray = varGetMarketDataJaggArray(strDirectory, varReportFileNameArray) 'Получаем MarketData.
    varJoinRowNumJaggArray = varGetJoinRowNumJaggArray(varMarketDataJaggArray, byteJoinDepth) 'Получаем точки соединения.
    varBaseFragmentsMarketDataJaggArray = varGetBaseFragmentsMarketDataJaggArray(varMarketDataJaggArray, varJoinRowNumJaggArray) 'Получаем осн. фрагменты соединения MarketData.
    varAvgAddFragmentsMarketDataJaggArray = varGetAvgAddFragmentsMarketDataJaggArray(varMarketDataJaggArray, varJoinRowNumJaggArray, strJoinMode) 'Получаем усредненные доп. фрагменты соединения MarketData.
    intAggSizeBaseFragmentsMarketDataJaggArray = intGetAggSizeJaggArray(varBaseFragmentsMarketDataJaggArray) 'Получаем агрегированный размер массива массивов осн. фрагментов.
    intAggSizeAvgAddFragmentsMarketDataJaggArray = intGetAggSizeJaggArray(varAvgAddFragmentsMarketDataJaggArray) 'Получаем агрегированный размер массива массивов усредненных доп. фрагментов.
    ReDim varJoinMarketDataArray(intAggSizeBaseFragmentsMarketDataJaggArray + intAggSizeAvgAddFragmentsMarketDataJaggArray, 1 To 6)
    For i = 1 To UBound(varBaseFragmentsMarketDataJaggArray) - 1 ' От первого массива осн. фраментов до предпоследнего т.к. осн. всегда на один больше.
        byteTrigger = TargetAddition(varJoinMarketDataArray, varBaseFragmentsMarketDataJaggArray(i)) 'Добавляем осн. фрагмент.
        byteTrigger = TargetAddition(varJoinMarketDataArray, varAvgAddFragmentsMarketDataJaggArray(i)) 'Добавляем доп. фрагмент (соблюдаем очередность: осн., доп., осн., доп., и т.д.).
    Next i
    byteTrigger = TargetAddition(varJoinMarketDataArray, varBaseFragmentsMarketDataJaggArray(UBound(varBaseFragmentsMarketDataJaggArray))) 'Добавляем последний осн. фрагмент.
    varGetJoinMarketDataArray = varJoinMarketDataArray
End Function

'Функция дает соединенный массив MarketData, если параметр JoinDepth=0.
Private Function varGetJoinMarketDataArray0(strDirectory As String, _
                                            varReportFileNameArray As Variant, _
                                            byteJoinDepth As Byte)
    Dim i As Integer, _
        varMarketDataJaggArray As Variant, _
        varJoinRowNumJaggArray As Variant, _
        varBaseFragmentsMarketDataJaggArray As Variant, _
        intAggSizeBaseFragmentsMarketDataJaggArray As Integer, _
        varJoinMarketDataArray() As Variant
    varMarketDataJaggArray = varGetMarketDataJaggArray(strDirectory, varReportFileNameArray) 'Получаем MarketData.
    varJoinRowNumJaggArray = varGetJoinRowNumJaggArray(varMarketDataJaggArray, byteJoinDepth) 'Получаем точки соединения.
    varBaseFragmentsMarketDataJaggArray = varGetBaseFragmentsMarketDataJaggArray(varMarketDataJaggArray, varJoinRowNumJaggArray) 'Получаем осн. фрагменты соединения MarketData.
    intAggSizeBaseFragmentsMarketDataJaggArray = intGetAggSizeJaggArray(varBaseFragmentsMarketDataJaggArray) 'Получаем агрегированный размер массива массивов осн. фрагментов.
    ReDim varJoinMarketDataArray(intAggSizeBaseFragmentsMarketDataJaggArray, 1 To 6)
    For i = 1 To UBound(varBaseFragmentsMarketDataJaggArray) ' От первого массива осн. фраментов до последнего.
        byteTrigger = TargetAddition(varJoinMarketDataArray, varBaseFragmentsMarketDataJaggArray(i)) 'Добавляем осн. фрагмент.
    Next i
    varGetJoinMarketDataArray0 = varJoinMarketDataArray
End Function

'Функция дает номер строки начальной даты.
Private Function intGetBegDateRowNum(varJoinMarketDataArray As Variant, _
                                     strBegDate As String)
    Dim i As Integer
    For i = 1 To UBound(varJoinMarketDataArray, 1)
        If DateValue(varJoinMarketDataArray(i, 1)) >= DateValue(strBegDate) Then 'Включает примагничивание к ближайшей большей дате.
            intGetBegDateRowNum = i
            Exit Function
        End If
    Next i
End Function

'Функция дает номер строки конечной даты.
Private Function intGetEndDateRowNum(varJoinMarketDataArray As Variant, _
                                     strEndDate As String)
    Dim i As Integer
    For i = UBound(varJoinMarketDataArray, 1) To 1 Step -1
        If DateValue(varJoinMarketDataArray(i, 1)) <= DateValue(strEndDate) Then 'Включает примагничивание к ближайшей меньшей дате.
            intGetEndDateRowNum = i
            Exit Function
        End If
    Next i
End Function

'Функция выводит соединенный массив MarketData.
Private Function OutputJoinMarketDataArray(varJoinMarketDataArray As Variant, _
                                           intRow_TradeDate As Integer, intColumn_TradeDate As Integer, _
                                           intRow_Settle As Integer, intColumn_Settle As Integer, _
                                           intRow_Open As Integer, intColumn_Open As Integer, _
                                           intRow_High As Integer, intColumn_High As Integer, _
                                           intRow_Low As Integer, intColumn_Low As Integer, _
                                           intRow_Close As Integer, intColumn_Close As Integer, _
                                           intBegDateRowNum As Integer, intEndDateRowNum As Integer)
    Dim i As Integer
    Const lngMaxQuad As Long = 10 'Сколько квадратов выводить.
    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("AutoHV").Range(Cells(intRow_TradeDate + 1, intColumn_TradeDate), Cells(intRow_Close + 1000, intColumn_Close)).ClearContents
    For i = intBegDateRowNum To intEndDateRowNum
        With Worksheets("AutoHV")
            .Cells(intRow_TradeDate + i - (intBegDateRowNum - 1), intColumn_TradeDate).Value = varJoinMarketDataArray(i, 1)
            .Cells(intRow_Settle + 1 + i - (intBegDateRowNum - 1), intColumn_Settle).Value = varJoinMarketDataArray(i, 2)
            .Cells(intRow_Open + 1 + i - (intBegDateRowNum - 1), intColumn_Open).Value = varJoinMarketDataArray(i, 3)
            .Cells(intRow_High + 1 + i - (intBegDateRowNum - 1), intColumn_High).Value = varJoinMarketDataArray(i, 4)
            .Cells(intRow_Low + 1 + i - (intBegDateRowNum - 1), intColumn_Low).Value = varJoinMarketDataArray(i, 5)
            .Cells(intRow_Close + 1 + i - (intBegDateRowNum - 1), intColumn_Close).Value = varJoinMarketDataArray(i, 6)
        End With
        Application.StatusBar = "Выполнено: " & Int(100 * (i - (intBegDateRowNum - 1)) / ((intEndDateRowNum - intBegDateRowNum) + 1)) & "%" _
        & String(CLng(lngMaxQuad * (i - (intBegDateRowNum - 1)) / ((intEndDateRowNum - intBegDateRowNum) + 1)), ChrW(9632)) _
        & String(lngMaxQuad - CLng(lngMaxQuad * (i - (intBegDateRowNum - 1)) / ((intEndDateRowNum - intBegDateRowNum) + 1)), ChrW(9633))
        DoEvents 'Перерисовка формы.
    Next i
    Application.StatusBar = False 'Сбрасываем значение статусной строки.
End Function

'Схема относительных изменений типа текущий/предыдущий.
'P0dte - добавочная строка (BegRowNum)
'P1dte_P1dte/P0dte
'Pndte_Pndte/P1dte_Pndte/P1dte
'P1____P1/Pndte____P1/Pndte____P1/Pndte
'P2________________P2/P1_______P2/P1_P2/P1
'P3____________________________P3/P2_P3/P2_P3/P2
'P4__________________________________P4/P3_P4/P3
'Pn________________________________________Pn/P4 (EndRowNum)
'Схема относительных изменений типа текущий/текущий.
'P1dte_P1dte/P1dte (BegRowNum)
'Pndte_Pndte/Pndte_Pndte/Pndte
'P1____P1/P1_______P1/P1_______P1/P1
'P2________________P2/P2_______P2/P2_P2/P2
'P3____________________________P3/P3_P3/P3_P3/P3
'P4__________________________________P4/P4_P4/P4
'Pn________________________________________Pn/Pn (EndRowNum)
'Функция дает массив лог. относ. цен тек./пред. ln(Pcur/Pprev).
Private Function varGetLnCurPrevArray(varJoinMarketDataArray As Variant, _
                                      byteDim1 As Byte, _
                                      byteDim2 As Byte, _
                                      byteFrequency As Byte, _
                                      varDTE As Variant, _
                                      intBegRowNum As Integer, _
                                      intEndRowNum As Integer) As Variant
    Dim i As Integer, _
        intCorrBegRowNum As Integer, _
        intLnCounter As Integer, _
        dblLnCurPrevArray() As Double
    intCorrBegRowNum = intBegRowNum - byteFrequency * varDTE 'Номер строки первого расчетного значения (стартовая точка).
    ReDim dblLnCurPrevArray(Int((intEndRowNum - intCorrBegRowNum) / byteFrequency)) 'Кол-во лог. относ. цен тек./пред. меньше кол-ва значений на 1.
    intLnCounter = 1
    For i = intCorrBegRowNum + byteFrequency To intEndRowNum Step byteFrequency 'От второго до последнего.
        dblLnCurPrevArray(intLnCounter) = Log(varJoinMarketDataArray(i, byteDim1) / varJoinMarketDataArray(i - byteFrequency, byteDim2))
        intLnCounter = intLnCounter + 1
    Next i
    varGetLnCurPrevArray = dblLnCurPrevArray
End Function

'Функция дает массив лог. относ. цен тек./тек. ln(Pcur/Pcur).
Private Function varGetLnCurCurArray(varJoinMarketDataArray As Variant, _
                                     byteDim1 As Byte, _
                                     byteDim2 As Byte, _
                                     byteFrequency As Byte, _
                                     varDTE As Variant, _
                                     intBegRowNum As Integer, _
                                     intEndRowNum As Integer) As Variant
    Dim i As Integer, _
        intCorrBegRowNum As Integer, _
        intLnCounter As Integer, _
        dblLnCurCurArray() As Double
    intCorrBegRowNum = (intBegRowNum - byteFrequency * varDTE) + byteFrequency 'Номер строки первого расчетного значения (стартовая точка).
    ReDim dblLnCurCurArray(Int((intEndRowNum - intCorrBegRowNum) / byteFrequency) + 1) 'Кол-во лог. относ. цен тек./тек. соот. кол-ву значений.
    intLnCounter = 1
    For i = intCorrBegRowNum To intEndRowNum Step byteFrequency 'От первого до последнего.
        dblLnCurCurArray(intLnCounter) = Log(varJoinMarketDataArray(i, byteDim1) / varJoinMarketDataArray(i, byteDim2))
        intLnCounter = intLnCounter + 1
    Next i
    varGetLnCurCurArray = dblLnCurCurArray
End Function

'Функция дает фрагмент многомерного массива по заданному измерению (учитывается частота).
Private Function varGetFragmentMultidimArray(varMultidimArray As Variant, _
                                             byteDim As Byte, _
                                             byteFrequency As Byte, _
                                             intBegRowNum As Integer, _
                                             intEndRowNum As Integer) As Variant
    Dim i As Integer, _
        intDateCounter As Integer, _
        varFragmentMultidimArray() As Variant
    ReDim varFragmentMultidimArray(Int((intEndRowNum - intBegRowNum) / byteFrequency) + 1)
    intDateCounter = 1
    For i = intBegRowNum To intEndRowNum Step byteFrequency
        varFragmentMultidimArray(intDateCounter) = varMultidimArray(i, byteDim)
        intDateCounter = intDateCounter + 1
    Next i
    varGetFragmentMultidimArray = varFragmentMultidimArray
End Function

'Функция дает фрагмент массива.
Private Function varGetFragmentArray(varArray As Variant, _
                                     intBegRowNum As Integer, _
                                     intEndRowNum As Integer) As Variant
    Dim i As Integer, _
        varFragmentArray() As Variant
    ReDim varFragmentArray((intEndRowNum - intBegRowNum) + 1)
    For i = intBegRowNum To intEndRowNum
        varFragmentArray(i - (intBegRowNum - 1)) = varArray(i)
    Next i
    varGetFragmentArray = varFragmentArray
End Function

'Функция дает массив дисперсий.
Private Function varGetVarianceArray(varLnArray As Variant, _
                                     varDTE As Variant) As Variant
    Dim i As Integer, _
        varFragmentArray As Variant, _
        varVarianceArray() As Variant, _
        f As WorksheetFunction
    Set f = Application.WorksheetFunction
    ReDim varVarianceArray((UBound(varLnArray) - varDTE) + 1)
    For i = 1 To UBound(varVarianceArray)
        varFragmentArray = varGetFragmentArray(varLnArray, i, (varDTE + i) - 1)
        varVarianceArray(i) = f.var(varFragmentArray)
    Next i
    varGetVarianceArray = varVarianceArray
End Function

'Функция дает массив волатильностей.
Private Function varGetVolaArray(varVarianceArray As Variant, _
                                 byteFrequency As Byte, _
                                 varLOT As Variant) As Variant
    Dim i As Integer, _
        varVolaArray() As Variant
    ReDim varVolaArray(UBound(varVarianceArray)) 'Кол-во вол. соот. кол-ву дисп.
    For i = 1 To UBound(varVolaArray)
        varVolaArray(i) = Sqr(varLOT / byteFrequency) * Sqr(varVarianceArray(i))
    Next i
    varGetVolaArray = varVolaArray
End Function

'Функция дает ClosetoClose массив дисперсий.
Private Function varGetClosetoCloseVarianceArray(varJoinMarketDataArray As Variant, _
                                                 byteFrequency As Byte, _
                                                 varDTE As Variant, _
                                                 intBegDateRowNum As Integer, _
                                                 intEndDateRowNum As Integer) As Variant
    Dim varLnCcCpArray As Variant
    varLnCcCpArray = varGetLnCurPrevArray(varJoinMarketDataArray, 6, 6, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Cc/Cp).
    varGetClosetoCloseVarianceArray = varGetVarianceArray(varLnCcCpArray, varDTE)
End Function

'Функция дает ClosetoClose массив массивов.
Private Function varGetClosetoCloseJaggArray(varJoinMarketDataArray As Variant, _
                                             byteFrequency As Byte, _
                                             varLOT As Variant, _
                                             intBegDateRowNum As Integer, _
                                             intEndDateRowNum As Integer) As Variant
    Dim varJaggArray(9) As Variant
    varJaggArray(1) = varGetFragmentMultidimArray(varJoinMarketDataArray, 1, byteFrequency, intBegDateRowNum, intEndDateRowNum)  'Получаем массив дат.
    varJaggArray(2) = varGetClosetoCloseVarianceArray(varJoinMarketDataArray, byteFrequency, 100, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 100.
    varJaggArray(3) = varGetVolaArray(varJaggArray(2), byteFrequency, varLOT) 'Получаем массив вол 100.
    varJaggArray(4) = varGetClosetoCloseVarianceArray(varJoinMarketDataArray, byteFrequency, 50, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 50.
    varJaggArray(5) = varGetVolaArray(varJaggArray(4), byteFrequency, varLOT) 'Получаем массив вол 50.
    varJaggArray(6) = varGetClosetoCloseVarianceArray(varJoinMarketDataArray, byteFrequency, 20, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 20.
    varJaggArray(7) = varGetVolaArray(varJaggArray(6), byteFrequency, varLOT) 'Получаем массив вол 20.
    varJaggArray(8) = varGetClosetoCloseVarianceArray(varJoinMarketDataArray, byteFrequency, 10, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 10.
    varJaggArray(9) = varGetVolaArray(varJaggArray(8), byteFrequency, varLOT) 'Получаем массив вол 20.
    varGetClosetoCloseJaggArray = varJaggArray
End Function

'Функция дает Parkinson массив дисперсий.
Private Function varGetParkinsonVarianceArray(varJoinMarketDataArray As Variant, _
                                              byteFrequency As Byte, _
                                              varDTE As Variant, _
                                              intBegDateRowNum As Integer, _
                                              intEndDateRowNum As Integer) As Variant
    Dim i As Integer, _
        varLnHcLcArray As Variant, _
        varFragmentLnHcLcArray As Variant, _
        varParkinsonVarianceArray() As Variant, _
        f As WorksheetFunction
    Set f = Application.WorksheetFunction
    varLnHcLcArray = varGetLnCurCurArray(varJoinMarketDataArray, 4, 5, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Hc/Lc).
    ReDim varParkinsonVarianceArray((UBound(varLnHcLcArray) - varDTE) + 1)
    For i = 1 To UBound(varParkinsonVarianceArray)
        varFragmentLnHcLcArray = varGetFragmentArray(varLnHcLcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Hc/Lc).
        'Операторы с одинаковым приоритетом выполняются в порядке слева направо в том порядке, в котором они появляются в выражении.
        varParkinsonVarianceArray(i) = 1 / (4 * Log(2)) * f.SumSq(varFragmentLnHcLcArray) / varDTE 'Получаем дисперсию.
    Next i
    varGetParkinsonVarianceArray = varParkinsonVarianceArray
End Function

'Функция дает Parkinson массив массивов.
Private Function varGetParkinsonJaggArray(varJoinMarketDataArray As Variant, _
                                          byteFrequency As Byte, _
                                          varLOT As Variant, _
                                          intBegDateRowNum As Integer, _
                                          intEndDateRowNum As Integer) As Variant
    Dim varJaggArray(9) As Variant
    varJaggArray(1) = varGetFragmentMultidimArray(varJoinMarketDataArray, 1, byteFrequency, intBegDateRowNum, intEndDateRowNum)  'Получаем массив дат.
    varJaggArray(2) = varGetParkinsonVarianceArray(varJoinMarketDataArray, byteFrequency, 100, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 100.
    varJaggArray(3) = varGetVolaArray(varJaggArray(2), byteFrequency, varLOT) 'Получаем массив вол 100.
    varJaggArray(4) = varGetParkinsonVarianceArray(varJoinMarketDataArray, byteFrequency, 50, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 50.
    varJaggArray(5) = varGetVolaArray(varJaggArray(4), byteFrequency, varLOT) 'Получаем массив вол 50.
    varJaggArray(6) = varGetParkinsonVarianceArray(varJoinMarketDataArray, byteFrequency, 20, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 20.
    varJaggArray(7) = varGetVolaArray(varJaggArray(6), byteFrequency, varLOT) 'Получаем массив вол 20.
    varJaggArray(8) = varGetParkinsonVarianceArray(varJoinMarketDataArray, byteFrequency, 10, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 10.
    varJaggArray(9) = varGetVolaArray(varJaggArray(8), byteFrequency, varLOT) 'Получаем массив вол 20.
    varGetParkinsonJaggArray = varJaggArray
End Function

'Функция дает GarmanKlass массив дисперсий.
Private Function varGetGarmanKlassVarianceArray(varJoinMarketDataArray As Variant, _
                                                byteFrequency As Byte, _
                                                varDTE As Variant, _
                                                intBegDateRowNum As Integer, _
                                                intEndDateRowNum As Integer) As Variant
    Dim i As Integer, j As Integer, _
        varLnHcLcArray As Variant, _
        varLnCcOcArray As Variant, _
        varFragmentLnHcLcArray As Variant, _
        varFragmentLnCcOcArray As Variant, _
        varFragmentCompArray() As Variant, _
        varGarmanKlassVarianceArray() As Variant, _
        f As WorksheetFunction
    Set f = Application.WorksheetFunction
    varLnHcLcArray = varGetLnCurCurArray(varJoinMarketDataArray, 4, 5, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Hc/Lc).
    varLnCcOcArray = varGetLnCurCurArray(varJoinMarketDataArray, 6, 3, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Cc/Oc).
    ReDim varGarmanKlassVarianceArray((UBound(varLnHcLcArray) - varDTE) + 1)
    ReDim varFragmentCompArray(varDTE)
    For i = 1 To UBound(varGarmanKlassVarianceArray)
        varFragmentLnHcLcArray = varGetFragmentArray(varLnHcLcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Hc/Lc).
        varFragmentLnCcOcArray = varGetFragmentArray(varLnCcOcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Cc/Oc).
        For j = 1 To UBound(varFragmentCompArray) 'Получаем составной фрагмент.
            varFragmentCompArray(j) = 0.5 * varFragmentLnHcLcArray(j) ^ 2 - (2 * Log(2) - 1) * varFragmentLnCcOcArray(j) ^ 2
        Next j
        varGarmanKlassVarianceArray(i) = f.Sum(varFragmentCompArray) / varDTE 'Получаем дисперсию.
    Next i
    varGetGarmanKlassVarianceArray = varGarmanKlassVarianceArray
End Function

'Функция дает GarmanKlass массив массивов.
Private Function varGetGarmanKlassJaggArray(varJoinMarketDataArray As Variant, _
                                            byteFrequency As Byte, _
                                            varLOT As Variant, _
                                            intBegDateRowNum As Integer, _
                                            intEndDateRowNum As Integer) As Variant
    Dim varJaggArray(9) As Variant
    varJaggArray(1) = varGetFragmentMultidimArray(varJoinMarketDataArray, 1, byteFrequency, intBegDateRowNum, intEndDateRowNum)  'Получаем массив дат.
    varJaggArray(2) = varGetGarmanKlassVarianceArray(varJoinMarketDataArray, byteFrequency, 100, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 100.
    varJaggArray(3) = varGetVolaArray(varJaggArray(2), byteFrequency, varLOT) 'Получаем массив вол 100.
    varJaggArray(4) = varGetGarmanKlassVarianceArray(varJoinMarketDataArray, byteFrequency, 50, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 50.
    varJaggArray(5) = varGetVolaArray(varJaggArray(4), byteFrequency, varLOT) 'Получаем массив вол 50.
    varJaggArray(6) = varGetGarmanKlassVarianceArray(varJoinMarketDataArray, byteFrequency, 20, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 20.
    varJaggArray(7) = varGetVolaArray(varJaggArray(6), byteFrequency, varLOT) 'Получаем массив вол 20.
    varJaggArray(8) = varGetGarmanKlassVarianceArray(varJoinMarketDataArray, byteFrequency, 10, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 10.
    varJaggArray(9) = varGetVolaArray(varJaggArray(8), byteFrequency, varLOT) 'Получаем массив вол 20.
    varGetGarmanKlassJaggArray = varJaggArray
End Function

'Функция дает RogersSatchell массив дисперсий.
Private Function varGetRogersSatchellVarianceArray(varJoinMarketDataArray As Variant, _
                                                   byteFrequency As Byte, _
                                                   varDTE As Variant, _
                                                   intBegDateRowNum As Integer, _
                                                   intEndDateRowNum As Integer) As Variant
    Dim i As Integer, j As Integer, _
        varLnHcCcArray As Variant, _
        varLnHcOcArray As Variant, _
        varLnLcCcArray As Variant, _
        varLnLcOcArray As Variant, _
        varFragmentLnHcCcArray As Variant, _
        varFragmentLnHcOcArray As Variant, _
        varFragmentLnLcCcArray As Variant, _
        varFragmentLnLcOcArray As Variant, _
        varFragmentCompArray() As Variant, _
        varRogersSatchellVarianceArray() As Variant, _
        f As WorksheetFunction
    Set f = Application.WorksheetFunction
    varLnHcCcArray = varGetLnCurCurArray(varJoinMarketDataArray, 4, 6, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Hc/Cc).
    varLnHcOcArray = varGetLnCurCurArray(varJoinMarketDataArray, 4, 3, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Hc/Oc).
    varLnLcCcArray = varGetLnCurCurArray(varJoinMarketDataArray, 5, 6, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Lc/Cc).
    varLnLcOcArray = varGetLnCurCurArray(varJoinMarketDataArray, 5, 3, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Lc/Oc).
    ReDim varRogersSatchellVarianceArray((UBound(varLnHcCcArray) - varDTE) + 1)
    ReDim varFragmentCompArray(varDTE)
    For i = 1 To UBound(varRogersSatchellVarianceArray)
        varFragmentLnHcCcArray = varGetFragmentArray(varLnHcCcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Hc/Cc).
        varFragmentLnHcOcArray = varGetFragmentArray(varLnHcOcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Hc/Oc).
        varFragmentLnLcCcArray = varGetFragmentArray(varLnLcCcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Lc/Cc).
        varFragmentLnLcOcArray = varGetFragmentArray(varLnLcOcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Lc/Oc).
        For j = 1 To UBound(varFragmentCompArray) 'Получаем составной фрагмент.
            varFragmentCompArray(j) = varFragmentLnHcCcArray(j) * varFragmentLnHcOcArray(j) + varFragmentLnLcCcArray(j) * varFragmentLnLcOcArray(j)
        Next j
        varRogersSatchellVarianceArray(i) = f.Average(varFragmentCompArray) 'Получаем дисперсию.
    Next i
    varGetRogersSatchellVarianceArray = varRogersSatchellVarianceArray
End Function

'Функция дает RogersSatchell массив массивов.
Private Function varGetRogersSatchellJaggArray(varJoinMarketDataArray As Variant, _
                                               byteFrequency As Byte, _
                                               varLOT As Variant, _
                                               intBegDateRowNum As Integer, _
                                               intEndDateRowNum As Integer) As Variant
    Dim varJaggArray(9) As Variant
    varJaggArray(1) = varGetFragmentMultidimArray(varJoinMarketDataArray, 1, byteFrequency, intBegDateRowNum, intEndDateRowNum)  'Получаем массив дат.
    varJaggArray(2) = varGetRogersSatchellVarianceArray(varJoinMarketDataArray, byteFrequency, 100, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 100.
    varJaggArray(3) = varGetVolaArray(varJaggArray(2), byteFrequency, varLOT) 'Получаем массив вол 100.
    varJaggArray(4) = varGetRogersSatchellVarianceArray(varJoinMarketDataArray, byteFrequency, 50, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 50.
    varJaggArray(5) = varGetVolaArray(varJaggArray(4), byteFrequency, varLOT) 'Получаем массив вол 50.
    varJaggArray(6) = varGetRogersSatchellVarianceArray(varJoinMarketDataArray, byteFrequency, 20, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 20.
    varJaggArray(7) = varGetVolaArray(varJaggArray(6), byteFrequency, varLOT) 'Получаем массив вол 20.
    varJaggArray(8) = varGetRogersSatchellVarianceArray(varJoinMarketDataArray, byteFrequency, 10, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 10.
    varJaggArray(9) = varGetVolaArray(varJaggArray(8), byteFrequency, varLOT) 'Получаем массив вол 20.
    varGetRogersSatchellJaggArray = varJaggArray
End Function

'Функция дает GarmanKlassYangZhangExtension массив дисперсий.
Private Function varGetGarmanKlassYangZhangExtensionVarianceArray(varJoinMarketDataArray As Variant, _
                                                                  byteFrequency As Byte, _
                                                                  varDTE As Variant, _
                                                                  intBegDateRowNum As Integer, _
                                                                  intEndDateRowNum As Integer) As Variant
    Dim i As Integer, j As Integer, _
        varLnOcCpArray As Variant, _
        varLnHcLcArray As Variant, _
        varLnCcOcArray As Variant, _
        varFragmentLnOcCpArray As Variant, _
        varFragmentLnHcLcArray As Variant, _
        varFragmentLnCcOcArray As Variant, _
        varFragmentCompArray() As Variant, _
        varGarmanKlassYangZhangExtensionVarianceArray() As Variant, _
        f As WorksheetFunction
    Set f = Application.WorksheetFunction
    varLnOcCpArray = varGetLnCurPrevArray(varJoinMarketDataArray, 3, 6, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Oc/Cp).
    varLnHcLcArray = varGetLnCurCurArray(varJoinMarketDataArray, 4, 5, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Hc/Lc).
    varLnCcOcArray = varGetLnCurCurArray(varJoinMarketDataArray, 6, 3, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Cc/Oc).
    ReDim varGarmanKlassYangZhangExtensionVarianceArray((UBound(varLnOcCpArray) - varDTE) + 1)
    ReDim varFragmentCompArray(varDTE)
    For i = 1 To UBound(varGarmanKlassYangZhangExtensionVarianceArray)
        varFragmentLnOcCpArray = varGetFragmentArray(varLnOcCpArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Oc/Cp).
        varFragmentLnHcLcArray = varGetFragmentArray(varLnHcLcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Hc/Lc).
        varFragmentLnCcOcArray = varGetFragmentArray(varLnCcOcArray, i, (varDTE + i) - 1) 'Получаем фрагмент массива Ln(Cc/Oc).
        For j = 1 To UBound(varFragmentCompArray) 'Получаем составной фрагмент.
            varFragmentCompArray(j) = varFragmentLnOcCpArray(j) ^ 2 + 0.5 * varFragmentLnHcLcArray(j) ^ 2 - (2 * Log(2) - 1) * varFragmentLnCcOcArray(j) ^ 2
        Next j
        varGarmanKlassYangZhangExtensionVarianceArray(i) = f.Sum(varFragmentCompArray) / varDTE 'Получаем дисперсию.
    Next i
    varGetGarmanKlassYangZhangExtensionVarianceArray = varGarmanKlassYangZhangExtensionVarianceArray
End Function

'Функция дает GarmanKlassYangZhangExtension массив массивов.
Private Function varGetGarmanKlassYangZhangExtensionJaggArray(varJoinMarketDataArray As Variant, _
                                                              byteFrequency As Byte, _
                                                              varLOT As Variant, _
                                                              intBegDateRowNum As Integer, _
                                                              intEndDateRowNum As Integer) As Variant
    Dim varJaggArray(9) As Variant
    varJaggArray(1) = varGetFragmentMultidimArray(varJoinMarketDataArray, 1, byteFrequency, intBegDateRowNum, intEndDateRowNum)  'Получаем массив дат.
    varJaggArray(2) = varGetGarmanKlassYangZhangExtensionVarianceArray(varJoinMarketDataArray, byteFrequency, 100, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 100.
    varJaggArray(3) = varGetVolaArray(varJaggArray(2), byteFrequency, varLOT) 'Получаем массив вол 100.
    varJaggArray(4) = varGetGarmanKlassYangZhangExtensionVarianceArray(varJoinMarketDataArray, byteFrequency, 50, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 50.
    varJaggArray(5) = varGetVolaArray(varJaggArray(4), byteFrequency, varLOT) 'Получаем массив вол 50.
    varJaggArray(6) = varGetGarmanKlassYangZhangExtensionVarianceArray(varJoinMarketDataArray, byteFrequency, 20, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 20.
    varJaggArray(7) = varGetVolaArray(varJaggArray(6), byteFrequency, varLOT) 'Получаем массив вол 20.
    varJaggArray(8) = varGetGarmanKlassYangZhangExtensionVarianceArray(varJoinMarketDataArray, byteFrequency, 10, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 10.
    varJaggArray(9) = varGetVolaArray(varJaggArray(8), byteFrequency, varLOT) 'Получаем массив вол 20.
    varGetGarmanKlassYangZhangExtensionJaggArray = varJaggArray
End Function

'Функция дает YangZhang массив дисперсий.
Private Function varGetYangZhangVarianceArray(varJoinMarketDataArray As Variant, _
                                              byteFrequency As Byte, _
                                              varDTE As Variant, _
                                              intBegDateRowNum As Integer, _
                                              intEndDateRowNum As Integer) As Variant
    Dim i As Integer, _
        dblCoeff As Double, _
        varLnOcCpArray As Variant, _
        varLnCcOcArray As Variant, _
        varOvernightVarianceArray As Variant, _
        varOpentoCloseVarianceArray As Variant, _
        varRogersSatchellVarianceArray As Variant, _
        varYangZhangVarianceArray() As Variant
    dblCoeff = 0.34 / (1.34 + (varDTE + 1) / (varDTE - 1)) 'Коэффициент.
    varLnOcCpArray = varGetLnCurPrevArray(varJoinMarketDataArray, 3, 6, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Oc/Cp).
    varLnCcOcArray = varGetLnCurCurArray(varJoinMarketDataArray, 6, 3, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем массив лог. относ. цен Ln(Cc/Oc).
    varOvernightVarianceArray = varGetVarianceArray(varLnOcCpArray, varDTE) 'Получаем массив дисперсий Overnight.
    varOpentoCloseVarianceArray = varGetVarianceArray(varLnCcOcArray, varDTE) 'Получаем массив дисперсий OpentoClose.
    varRogersSatchellVarianceArray = varGetRogersSatchellVarianceArray(varJoinMarketDataArray, byteFrequency, varDTE, intBegDateRowNum, intEndDateRowNum) 'Получаем RogersSatchell массив дисперсий.
    ReDim varYangZhangVarianceArray(UBound(varOvernightVarianceArray))
    For i = 1 To UBound(varYangZhangVarianceArray)
        varYangZhangVarianceArray(i) = varOvernightVarianceArray(i) + dblCoeff * varOpentoCloseVarianceArray(i) + (1 - dblCoeff) * varRogersSatchellVarianceArray(i)
    Next i
    varGetYangZhangVarianceArray = varYangZhangVarianceArray
End Function

'Функция дает YangZhang массив массивов.
Private Function varGetYangZhangJaggArray(varJoinMarketDataArray As Variant, _
                                          byteFrequency As Byte, _
                                          varLOT As Variant, _
                                          intBegDateRowNum As Integer, _
                                          intEndDateRowNum As Integer) As Variant
    Dim varJaggArray(9) As Variant
    varJaggArray(1) = varGetFragmentMultidimArray(varJoinMarketDataArray, 1, byteFrequency, intBegDateRowNum, intEndDateRowNum)  'Получаем массив дат.
    varJaggArray(2) = varGetYangZhangVarianceArray(varJoinMarketDataArray, byteFrequency, 100, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 100.
    varJaggArray(3) = varGetVolaArray(varJaggArray(2), byteFrequency, varLOT) 'Получаем массив вол 100.
    varJaggArray(4) = varGetYangZhangVarianceArray(varJoinMarketDataArray, byteFrequency, 50, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 50.
    varJaggArray(5) = varGetVolaArray(varJaggArray(4), byteFrequency, varLOT) 'Получаем массив вол 50.
    varJaggArray(6) = varGetYangZhangVarianceArray(varJoinMarketDataArray, byteFrequency, 20, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 20.
    varJaggArray(7) = varGetVolaArray(varJaggArray(6), byteFrequency, varLOT) 'Получаем массив вол 20.
    varJaggArray(8) = varGetYangZhangVarianceArray(varJoinMarketDataArray, byteFrequency, 10, intBegDateRowNum, intEndDateRowNum) 'Получаем массив дисп 10.
    varJaggArray(9) = varGetVolaArray(varJaggArray(8), byteFrequency, varLOT) 'Получаем массив вол 20.
    varGetYangZhangJaggArray = varJaggArray
End Function

'Функция выводит массивы Дата, Дисперсия, Волатильность.
Private Function OutputDateVarianceVola(varDateArray As Variant, _
                                        varVarianceArray As Variant, _
                                        varVolaArray As Variant, _
                                        intRow_TradeDate As Integer, _
                                        intColumn_TradeDate As Integer, _
                                        intRow_VarianceDTE As Integer, _
                                        intColumn_VarianceDTE As Integer, _
                                        intRow_VolaDTE As Integer, _
                                        intColumn_VolaDTE As Integer)
    Dim i As Integer
    Const lngMaxQuad As Long = 10 'Сколько квадратов выводить.
    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("AutoHV").Range(Cells(intRow_TradeDate + 1, intColumn_TradeDate), Cells(intRow_VolaDTE + 1000, intColumn_VolaDTE)).ClearContents
    For i = 1 To UBound(varDateArray)
        With Worksheets("AutoHV")
            .Cells(intRow_TradeDate + i, intColumn_TradeDate).Value = varDateArray(i)
            .Cells(intRow_VarianceDTE + i, intColumn_VarianceDTE).Value = varVarianceArray(i)
            .Cells(intRow_VolaDTE + i, intColumn_VolaDTE).Value = varVolaArray(i)
        End With
        Application.StatusBar = "Выполнено: " & Int(100 * i / UBound(varDateArray)) & "%" _
        & String(CLng(lngMaxQuad * i / UBound(varDateArray)), ChrW(9632)) _
        & String(lngMaxQuad - CLng(lngMaxQuad * i / UBound(varDateArray)), ChrW(9633))
        DoEvents 'Перерисовка формы.
    Next i
    Application.StatusBar = False 'Сбрасываем значение статусной строки.
End Function

'Функция выводит массив массивов Дата, 100502010Дисперсия, 100502010Волатильность.
Private Function OutputComboDateVarianceVola(varJaggArray As Variant, _
                                             intRow_TradeDate As Integer, _
                                             intColumn_TradeDate As Integer, _
                                             intRow_Variance100 As Integer, _
                                             intColumn_Variance100 As Integer, _
                                             intRow_Vola100 As Integer, _
                                             intColumn_Vola100 As Integer, _
                                             intRow_Variance50 As Integer, _
                                             intColumn_Variance50 As Integer, _
                                             intRow_Vola50 As Integer, _
                                             intColumn_Vola50 As Integer, _
                                             intRow_Variance20 As Integer, _
                                             intColumn_Variance20 As Integer, _
                                             intRow_Vola20 As Integer, _
                                             intColumn_Vola20 As Integer, _
                                             intRow_Variance10 As Integer, _
                                             intColumn_Variance10 As Integer, _
                                             intRow_Vola10 As Integer, _
                                             intColumn_Vola10 As Integer)
    Dim i As Integer
    Const lngMaxQuad As Long = 10 'Сколько квадратов выводить.
    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("AutoHV").Range(Cells(intRow_TradeDate + 1, intColumn_TradeDate), Cells(intRow_Vola10 + 1000, intColumn_Vola10)).ClearContents
    For i = 1 To UBound(varJaggArray(1))
        With Worksheets("AutoHV")
            .Cells(intRow_TradeDate + i, intColumn_TradeDate).Value = varJaggArray(1)(i)
            .Cells(intRow_Variance100 + i, intColumn_Variance100).Value = varJaggArray(2)(i)
            .Cells(intRow_Vola100 + i, intColumn_Vola100).Value = varJaggArray(3)(i)
            .Cells(intRow_Variance50 + i, intColumn_Variance50).Value = varJaggArray(4)(i)
            .Cells(intRow_Vola50 + i, intColumn_Vola50).Value = varJaggArray(5)(i)
            .Cells(intRow_Variance20 + i, intColumn_Variance20).Value = varJaggArray(6)(i)
            .Cells(intRow_Vola20 + i, intColumn_Vola20).Value = varJaggArray(7)(i)
            .Cells(intRow_Variance10 + i, intColumn_Variance10).Value = varJaggArray(8)(i)
            .Cells(intRow_Vola10 + i, intColumn_Vola10).Value = varJaggArray(9)(i)
        End With
        Application.StatusBar = "Выполнено: " & Int(100 * i / UBound(varJaggArray(1))) & "%" _
        & String(CLng(lngMaxQuad * i / UBound(varJaggArray(1))), ChrW(9632)) _
        & String(lngMaxQuad - CLng(lngMaxQuad * i / UBound(varJaggArray(1))), ChrW(9633))
        DoEvents 'Перерисовка формы.
    Next i
    Application.StatusBar = False 'Сбрасываем значение статусной строки.
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Название и формат полей CSV-файла Контракт_Код-ru (Например, SBRF-3.17_SRH7-ru).
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Дата      |Средневзвешенная|Расчетная|Цена первой|Максимальная|Минимальная|Цена     |Изменение|Объем    |Число |Объем  |Объем  |Объем   |Объем   |
'          |цена за лот     |цена     |сделки     |цена        |цена       |последней|         |последней|сделок|торгов,|торгов,|открытых|открытых|
'          |                |         |           |            |           |сделки   |         |сделки,  |      |руб.   |контр. |позиций,|позиций,|
'          |                |         |           |            |           |         |         |контр.   |      |       |       |руб.    |контр.  |
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'dd.mm.yyyy;                ;         ;           ;            ;           ;         ;         ;         ;      ;       ;       ;        ;        ;
'dd.mm.yyyy;                ;         ;           ;            ;           ;         ;         ;         ;      ;       ;       ;        ;        ;
'          1                2         3           4            5           6         7         8         9      10      11      12       13       14
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub AutoHV()
Attribute AutoHV.VB_Description = "Дает оценку HV."
Attribute AutoHV.VB_ProcData.VB_Invoke_Func = "h\n14"
    Dim i As Integer, j As Integer, _
        intRow_Directory As Integer, intColumn_Directory As Integer, _
        intRow_JoinMode As Integer, intColumn_JoinMode As Integer, _
        intRow_JoinDepth As Integer, intColumn_JoinDepth As Integer, _
        intRow_Estimator As Integer, intColumn_Estimator As Integer, _
        intRow_Frequency As Integer, intColumn_Frequency As Integer, _
        intRow_LOT As Integer, intColumn_LOT As Integer, _
        intRow_DTE As Integer, intColumn_DTE As Integer, _
        intRow_BegDate As Integer, intColumn_BegDate As Integer, _
        intRow_EndDate As Integer, intColumn_EndDate As Integer

    Dim intRow_TradeDate As Integer, intColumn_TradeDate As Integer, _
        intRow_Settle As Integer, intColumn_Settle As Integer, _
        intRow_Open As Integer, intColumn_Open As Integer, _
        intRow_High As Integer, intColumn_High As Integer, _
        intRow_Low As Integer, intColumn_Low As Integer, _
        intRow_Close As Integer, intColumn_Close As Integer, _
        intRow_VarianceDTE As Integer, intColumn_VarianceDTE As Integer, _
        intRow_VolaDTE As Integer, intColumn_VolaDTE As Integer, _
        intRow_Variance100 As Integer, intColumn_Variance100 As Integer, _
        intRow_Vola100 As Integer, intColumn_Vola100 As Integer, _
        intRow_Variance50 As Integer, intColumn_Variance50 As Integer, _
        intRow_Vola50 As Integer, intColumn_Vola50 As Integer, _
        intRow_Variance20 As Integer, intColumn_Variance20 As Integer, _
        intRow_Vola20 As Integer, intColumn_Vola20 As Integer, _
        intRow_Variance10 As Integer, intColumn_Variance10 As Integer, _
        intRow_Vola10 As Integer, intColumn_Vola10 As Integer

    Dim strDirectory As String, _
        strJoinMode As String, _
        byteJoinDepth As Byte, _
        strEstimator As String, _
        byteFrequency As Byte, _
        varLOT As Variant, _
        varDTE As Variant, _
        strBegDate As String, _
        strEndDate As String

    Dim varEstimatorAbbrArray As Variant, _
        varJoinModeAbbrArray As Variant, _
        intQtyReportFiles As Integer, _
        varReportFileNameArray As Variant, _
        varJoinMarketDataArray As Variant, _
        intBegDateRowNum As Integer, _
        intEndDateRowNum As Integer, _
        varDateArray As Variant, _
        varVarianceArray As Variant, _
        varVolaArray As Variant
    varEstimatorAbbrArray = Array("VCC", "VP", "VGK", "VRS", "VGKYZ", "VYZ")
    varJoinModeAbbrArray = Array("AM", "GM")
    'Получаем адреса ячеек входных параметров с панели пресетов.
    byteTrigger = GetAdressParam(intRow_Directory, intColumn_Directory, "Directory")
    byteTrigger = GetAdressParam(intRow_JoinMode, intColumn_JoinMode, "JoinMode")
    byteTrigger = GetAdressParam(intRow_JoinDepth, intColumn_JoinDepth, "JoinDepth")
    byteTrigger = GetAdressParam(intRow_Estimator, intColumn_Estimator, "Estimator")
    byteTrigger = GetAdressParam(intRow_Frequency, intColumn_Frequency, "Frequency")
    byteTrigger = GetAdressParam(intRow_LOT, intColumn_LOT, "LOT")
    byteTrigger = GetAdressParam(intRow_DTE, intColumn_DTE, "DTE")
    byteTrigger = GetAdressParam(intRow_BegDate, intColumn_BegDate, "BegDate")
    byteTrigger = GetAdressParam(intRow_EndDate, intColumn_EndDate, "EndDate")
    byteTrigger = GetAdressParam(intRow_TradeDate, intColumn_TradeDate, "TradeDate")
    byteTrigger = GetAdressParam(intRow_Settle, intColumn_Settle, "Settle")
    byteTrigger = GetAdressParam(intRow_Open, intColumn_Open, "Open")
    byteTrigger = GetAdressParam(intRow_High, intColumn_High, "High")
    byteTrigger = GetAdressParam(intRow_Low, intColumn_Low, "Low")
    byteTrigger = GetAdressParam(intRow_Close, intColumn_Close, "Close")
    byteTrigger = GetAdressParam(intRow_VarianceDTE, intColumn_VarianceDTE, "VarDTE")
    byteTrigger = GetAdressParam(intRow_VolaDTE, intColumn_VolaDTE, "VolaDTE")
    byteTrigger = GetAdressParam(intRow_Variance100, intColumn_Variance100, "Var100")
    byteTrigger = GetAdressParam(intRow_Vola100, intColumn_Vola100, "Vola100")
    byteTrigger = GetAdressParam(intRow_Variance50, intColumn_Variance50, "Var50")
    byteTrigger = GetAdressParam(intRow_Vola50, intColumn_Vola50, "Vola50")
    byteTrigger = GetAdressParam(intRow_Variance20, intColumn_Variance20, "Var20")
    byteTrigger = GetAdressParam(intRow_Vola20, intColumn_Vola20, "Vola20")
    byteTrigger = GetAdressParam(intRow_Variance10, intColumn_Variance10, "Var10")
    byteTrigger = GetAdressParam(intRow_Vola10, intColumn_Vola10, "Vola10")
    'Считываем значения с панели пресетов.
    strDirectory = Worksheets("AutoHV").Cells(intRow_Directory, intColumn_Directory + 1).Value
    strJoinMode = Worksheets("AutoHV").Cells(intRow_JoinMode, intColumn_JoinMode + 1).Value
    byteJoinDepth = Worksheets("AutoHV").Cells(intRow_JoinDepth, intColumn_JoinDepth + 1).Value
    strEstimator = Worksheets("AutoHV").Cells(intRow_Estimator, intColumn_Estimator + 1).Value
    byteFrequency = Worksheets("AutoHV").Cells(intRow_Frequency, intColumn_Frequency + 1).Value
    varLOT = Worksheets("AutoHV").Cells(intRow_LOT, intColumn_LOT + 1).Value
    varDTE = Worksheets("AutoHV").Cells(intRow_DTE, intColumn_DTE + 1).Value
    strBegDate = Worksheets("AutoHV").Cells(intRow_BegDate, intColumn_BegDate + 1).Value
    strEndDate = Worksheets("AutoHV").Cells(intRow_EndDate, intColumn_EndDate + 1).Value
   
    If strBegDate = "" And strEndDate = "" Then 'Условие для полной очистки рабочего пространства.
        Worksheets("AutoHV").Range(Cells(intRow_TradeDate + 1, intColumn_TradeDate), Cells(intRow_Vola10 + 1000, intColumn_Vola10)).ClearContents
    ElseIf strBegDate = "" Then 'Проверяем параметр BegDate.
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Не указана начальная дата." _
        & vbNewLine & "Добавьте значение параметра BegDate и повторите расчет HV." _
        & vbNewLine & "Необходимое условие - наличие параметров BegDate и EndDate.")
    ElseIf IsDate(strBegDate) = False Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Начальная дата некорректна." _
        & vbNewLine & "Исправьте значение параметра BegDate и повторите расчет HV.")
    ElseIf strEndDate = "" Then 'Проверяем параметр EndDate.
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Не указана конечная дата." _
        & vbNewLine & "Добавьте значение параметра EndDate и повторите расчет HV." _
        & vbNewLine & "Необходимое условие - наличие параметров BegDate и EndDate.")
    ElseIf IsDate(strEndDate) = False Then
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Конечная дата некорректна." _
        & vbNewLine & "Исправьте значение параметра EndDate и повторите расчет HV.")
    ElseIf strDirectory = "" Then 'Проверяем параметр Directory.
        MsgBox ("Оба-на!!!" _
        & vbNewLine & "Не указан путь к папке с файл-отчетами." _
        & vbNewLine & "Добавьте значение параметра Directory и повторите расчет HV.")
    Else
        If byteJoinDepth > 0 Then
            If strJoinMode = "" Then 'Проверяем параметр JoinMode.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Не указан режим соединения данных." _
                & vbNewLine & "Добавьте значение параметра JoinMode и повторите расчет HV." _
                & vbNewLine & "Доступные режимы соединения данных: AM, GM.")
                Exit Sub
            Else
                byteTrigger = boolHitTest(UCase(strJoinMode), varJoinModeAbbrArray) 'Тестируем значение параметра JoinMode на попадание.
                If byteTrigger = False Then 'Если нет попадания.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Режим соединения данных некорректен." _
                    & vbNewLine & "Исправьте значение параметра JoinMode и повторите расчет HV." _
                    & vbNewLine & "Доступные режимы соединения данных: AM, GM.")
                    Exit Sub
                End If
            End If
        End If
        byteTrigger = boolHitTest(UCase(strEstimator), varEstimatorAbbrArray) 'Тестируем значение параметра Estimator на попадание.
        If byteTrigger = False And strEstimator <> "" Then 'Если нет попадания и не пусто.
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "Тип оценки HV некорректен." _
            & vbNewLine & "Исправьте значение параметра Estimator и повторите расчет HV." _
            & vbNewLine & "Доступные типы оценки HV: Vcc,Vp,Vgk,Vrs,Vgkyz,Vyz.")
        Else 'Если есть попадание или пусто.
            If byteTrigger = True Then 'Если есть попадание.
                If byteFrequency = False Then 'Проверяем параметр Frequency.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Не указана частота измерений." _
                    & vbNewLine & "Добавьте значение параметра Frequency и повторите расчет HV.")
                    Exit Sub
                ElseIf IsEmpty(varLOT) = True Then 'Проверяем параметр LOT.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Не указана протяженность времени." _
                    & vbNewLine & "Добавьте значение параметра LOT и повторите расчет HV.")
                    Exit Sub
                End If
            End If
            intQtyReportFiles = intGetQtyReportFiles(strDirectory) 'Считаем количество файл-отчетов в папке назначения.
            If intQtyReportFiles > 0 Then 'Если в папке есть файл-отчеты.
                varReportFileNameArray = varGetReportFileNameArray(strDirectory, intQtyReportFiles) 'Создаем массив имен файл-отчетов.
                '<<<
                'THE NEW MOEX DATA SPEC DOES NOT REQUIRE SORTING

                'FUTREPORT = YYYY-TRADABLESYMBOL-TLD.EXTENSION
                'where
                'YYYY - settlement year, 4 symbols (0000 - 9999);
                'TRADABLESYMBOL = CMY
                'where
                'C – code of the underlying asset, 2 symbols;
                'M – settlement month (for options it also determines contract type – call or put), 1 symbol
                'where
                'M: Jan = F, Feb = G, Mar = H, Apr = J, May = K, Jun = M,
                'Jul = N, Aug = Q, Sep = U, Oct = V, Nov = X, Dec = Z;
                'Y – settlement year, 1 symbol (0-9).;
                'TLD - top-level domain (ru);
                'EXTENSION - file extension (csv).
                'Example of a report name: 2012-RIZ2-ru.csv, 2022-RIZ2-ru.csv
                'The leading year (YYYY) is added to eliminate confusion associated with the ten-year cycle.

                'DEPRECATED
                'varReportFileNameArray = varMainSortReportFileNameArray(varReportFileNameArray)
                '>>>
                If byteJoinDepth = 0 Then
                    varJoinMarketDataArray = varGetJoinMarketDataArray0(strDirectory, _
                                                                        varReportFileNameArray, _
                                                                        byteJoinDepth)
                Else
                    varJoinMarketDataArray = varGetJoinMarketDataArray(strDirectory, _
                                                                       varReportFileNameArray, _
                                                                       strJoinMode, _
                                                                       byteJoinDepth) 'Получаем соединенный массив MarketData.
                End If
                intBegDateRowNum = intGetBegDateRowNum(varJoinMarketDataArray, strBegDate) 'Получаем номер строки начальной даты.
                intEndDateRowNum = intGetEndDateRowNum(varJoinMarketDataArray, strEndDate) 'Получаем номер строки конечной даты.
                If strEstimator = "" Then 'Выводим соединенный массив MarketData.
                    byteTrigger = OutputJoinMarketDataArray(varJoinMarketDataArray, _
                                                            intRow_TradeDate, intColumn_TradeDate, _
                                                            intRow_Settle, intColumn_Settle, _
                                                            intRow_Open, intColumn_Open, _
                                                            intRow_High, intColumn_High, _
                                                            intRow_Low, intColumn_Low, _
                                                            intRow_Close, intColumn_Close, _
                                                            intBegDateRowNum, intEndDateRowNum)
                ElseIf UCase(strEstimator) = "VCC" Then 'Close-to-Close Volatility Estimator.
                    If IsEmpty(varDTE) = False Then
                        varDateArray = varGetFragmentMultidimArray(varJoinMarketDataArray, _
                                                                   1, _
                                                                   byteFrequency, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив дат.
                        varVarianceArray = varGetClosetoCloseVarianceArray(varJoinMarketDataArray, _
                                                                           byteFrequency, _
                                                                           varDTE, _
                                                                           intBegDateRowNum, _
                                                                           intEndDateRowNum) 'Получаем массив дисп.
                        varVolaArray = varGetVolaArray(varVarianceArray, byteFrequency, varLOT) 'Получаем массив вол.
                        byteTrigger = OutputDateVarianceVola(varDateArray, _
                                                             varVarianceArray, _
                                                             varVolaArray, _
                                                             intRow_TradeDate, intColumn_TradeDate, _
                                                             intRow_VarianceDTE, intColumn_VarianceDTE, _
                                                             intRow_VolaDTE, intColumn_VolaDTE) 'Выводим.
                    Else
                        varVolaArray = varGetClosetoCloseJaggArray(varJoinMarketDataArray, _
                                                                   byteFrequency, _
                                                                   varLOT, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив массивов дат, дисп., вол..
                        byteTrigger = OutputComboDateVarianceVola(varVolaArray, _
                                                                  intRow_TradeDate, intColumn_TradeDate, _
                                                                  intRow_Variance100, intColumn_Variance100, _
                                                                  intRow_Vola100, intColumn_Vola100, _
                                                                  intRow_Variance50, intColumn_Variance50, _
                                                                  intRow_Vola50, intColumn_Vola50, _
                                                                  intRow_Variance20, intColumn_Variance20, _
                                                                  intRow_Vola20, intColumn_Vola20, _
                                                                  intRow_Variance10, intColumn_Variance10, _
                                                                  intRow_Vola10, intColumn_Vola10) 'Выводим.
                    End If
                ElseIf UCase(strEstimator) = "VP" Then 'Parkinson Volatility Estimator.
                    If IsEmpty(varDTE) = False Then
                        varDateArray = varGetFragmentMultidimArray(varJoinMarketDataArray, _
                                                                   1, _
                                                                   byteFrequency, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив дат.
                        varVarianceArray = varGetParkinsonVarianceArray(varJoinMarketDataArray, _
                                                                        byteFrequency, _
                                                                        varDTE, _
                                                                        intBegDateRowNum, _
                                                                        intEndDateRowNum) 'Получаем массив дисп.
                        varVolaArray = varGetVolaArray(varVarianceArray, byteFrequency, varLOT) 'Получаем массив вол.
                        byteTrigger = OutputDateVarianceVola(varDateArray, _
                                                             varVarianceArray, _
                                                             varVolaArray, _
                                                             intRow_TradeDate, intColumn_TradeDate, _
                                                             intRow_VarianceDTE, intColumn_VarianceDTE, _
                                                             intRow_VolaDTE, intColumn_VolaDTE) 'Выводим.
                    Else
                        varVolaArray = varGetParkinsonJaggArray(varJoinMarketDataArray, _
                                                                byteFrequency, _
                                                                varLOT, _
                                                                intBegDateRowNum, _
                                                                intEndDateRowNum) 'Получаем массив массивов дат, дисп., вол..
                        byteTrigger = OutputComboDateVarianceVola(varVolaArray, _
                                                                  intRow_TradeDate, intColumn_TradeDate, _
                                                                  intRow_Variance100, intColumn_Variance100, _
                                                                  intRow_Vola100, intColumn_Vola100, _
                                                                  intRow_Variance50, intColumn_Variance50, _
                                                                  intRow_Vola50, intColumn_Vola50, _
                                                                  intRow_Variance20, intColumn_Variance20, _
                                                                  intRow_Vola20, intColumn_Vola20, _
                                                                  intRow_Variance10, intColumn_Variance10, _
                                                                  intRow_Vola10, intColumn_Vola10) 'Выводим.
                    End If
                ElseIf UCase(strEstimator) = "VGK" Then 'Garman-Klass Volatility Estimator.
                    If IsEmpty(varDTE) = False Then
                        varDateArray = varGetFragmentMultidimArray(varJoinMarketDataArray, _
                                                                   1, _
                                                                   byteFrequency, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив дат.
                        varVarianceArray = varGetGarmanKlassVarianceArray(varJoinMarketDataArray, _
                                                                          byteFrequency, _
                                                                          varDTE, _
                                                                          intBegDateRowNum, _
                                                                          intEndDateRowNum) 'Получаем массив дисп.
                        varVolaArray = varGetVolaArray(varVarianceArray, byteFrequency, varLOT) 'Получаем массив вол.
                        byteTrigger = OutputDateVarianceVola(varDateArray, _
                                                             varVarianceArray, _
                                                             varVolaArray, _
                                                             intRow_TradeDate, intColumn_TradeDate, _
                                                             intRow_VarianceDTE, intColumn_VarianceDTE, _
                                                             intRow_VolaDTE, intColumn_VolaDTE) 'Выводим.
                    Else
                        varVolaArray = varGetGarmanKlassJaggArray(varJoinMarketDataArray, _
                                                                  byteFrequency, _
                                                                  varLOT, _
                                                                  intBegDateRowNum, _
                                                                  intEndDateRowNum) 'Получаем массив массивов дат, дисп., вол..
                        byteTrigger = OutputComboDateVarianceVola(varVolaArray, _
                                                                  intRow_TradeDate, intColumn_TradeDate, _
                                                                  intRow_Variance100, intColumn_Variance100, _
                                                                  intRow_Vola100, intColumn_Vola100, _
                                                                  intRow_Variance50, intColumn_Variance50, _
                                                                  intRow_Vola50, intColumn_Vola50, _
                                                                  intRow_Variance20, intColumn_Variance20, _
                                                                  intRow_Vola20, intColumn_Vola20, _
                                                                  intRow_Variance10, intColumn_Variance10, _
                                                                  intRow_Vola10, intColumn_Vola10) 'Выводим.
                    End If
                ElseIf UCase(strEstimator) = "VRS" Then 'Rogers-Satchell Volatility Estimator.
                    If IsEmpty(varDTE) = False Then
                        varDateArray = varGetFragmentMultidimArray(varJoinMarketDataArray, _
                                                                   1, _
                                                                   byteFrequency, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив дат.
                        varVarianceArray = varGetRogersSatchellVarianceArray(varJoinMarketDataArray, _
                                                                             byteFrequency, _
                                                                             varDTE, _
                                                                             intBegDateRowNum, _
                                                                             intEndDateRowNum) 'Получаем массив дисп.
                        varVolaArray = varGetVolaArray(varVarianceArray, byteFrequency, varLOT) 'Получаем массив вол.
                        byteTrigger = OutputDateVarianceVola(varDateArray, _
                                                             varVarianceArray, _
                                                             varVolaArray, _
                                                             intRow_TradeDate, intColumn_TradeDate, _
                                                             intRow_VarianceDTE, intColumn_VarianceDTE, _
                                                             intRow_VolaDTE, intColumn_VolaDTE) 'Выводим.
                    Else
                        varVolaArray = varGetRogersSatchellJaggArray(varJoinMarketDataArray, _
                                                                     byteFrequency, _
                                                                     varLOT, _
                                                                     intBegDateRowNum, _
                                                                     intEndDateRowNum) 'Получаем массив массивов дат, дисп., вол..
                        byteTrigger = OutputComboDateVarianceVola(varVolaArray, _
                                                                  intRow_TradeDate, intColumn_TradeDate, _
                                                                  intRow_Variance100, intColumn_Variance100, _
                                                                  intRow_Vola100, intColumn_Vola100, _
                                                                  intRow_Variance50, intColumn_Variance50, _
                                                                  intRow_Vola50, intColumn_Vola50, _
                                                                  intRow_Variance20, intColumn_Variance20, _
                                                                  intRow_Vola20, intColumn_Vola20, _
                                                                  intRow_Variance10, intColumn_Variance10, _
                                                                  intRow_Vola10, intColumn_Vola10) 'Выводим.
                    End If
                ElseIf UCase(strEstimator) = "VGKYZ" Then 'Garman-Klass Yang-Zhang Extension Volatility Estimator.
                    If IsEmpty(varDTE) = False Then
                        varDateArray = varGetFragmentMultidimArray(varJoinMarketDataArray, _
                                                                   1, _
                                                                   byteFrequency, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив дат.
                        varVarianceArray = varGetGarmanKlassYangZhangExtensionVarianceArray(varJoinMarketDataArray, _
                                                                                            byteFrequency, _
                                                                                            varDTE, _
                                                                                            intBegDateRowNum, _
                                                                                            intEndDateRowNum) 'Получаем массив дисп.
                        varVolaArray = varGetVolaArray(varVarianceArray, byteFrequency, varLOT) 'Получаем массив вол.
                        byteTrigger = OutputDateVarianceVola(varDateArray, _
                                                             varVarianceArray, _
                                                             varVolaArray, _
                                                             intRow_TradeDate, intColumn_TradeDate, _
                                                             intRow_VarianceDTE, intColumn_VarianceDTE, _
                                                             intRow_VolaDTE, intColumn_VolaDTE) 'Выводим.
                    Else
                        varVolaArray = varGetGarmanKlassYangZhangExtensionJaggArray(varJoinMarketDataArray, _
                                                                                    byteFrequency, _
                                                                                    varLOT, _
                                                                                    intBegDateRowNum, _
                                                                                    intEndDateRowNum) 'Получаем массив массивов дат, дисп., вол..
                        byteTrigger = OutputComboDateVarianceVola(varVolaArray, _
                                                                  intRow_TradeDate, intColumn_TradeDate, _
                                                                  intRow_Variance100, intColumn_Variance100, _
                                                                  intRow_Vola100, intColumn_Vola100, _
                                                                  intRow_Variance50, intColumn_Variance50, _
                                                                  intRow_Vola50, intColumn_Vola50, _
                                                                  intRow_Variance20, intColumn_Variance20, _
                                                                  intRow_Vola20, intColumn_Vola20, _
                                                                  intRow_Variance10, intColumn_Variance10, _
                                                                  intRow_Vola10, intColumn_Vola10) 'Выводим.
                    End If
                ElseIf UCase(strEstimator) = "VYZ" Then 'Yang-Zhang Volatility Estimator.
                    If IsEmpty(varDTE) = False Then
                        varDateArray = varGetFragmentMultidimArray(varJoinMarketDataArray, _
                                                                   1, _
                                                                   byteFrequency, _
                                                                   intBegDateRowNum, _
                                                                   intEndDateRowNum) 'Получаем массив дат.
                        varVarianceArray = varGetYangZhangVarianceArray(varJoinMarketDataArray, _
                                                                        byteFrequency, _
                                                                        varDTE, _
                                                                        intBegDateRowNum, _
                                                                        intEndDateRowNum) 'Получаем массив дисп.
                        varVolaArray = varGetVolaArray(varVarianceArray, byteFrequency, varLOT) 'Получаем массив вол.
                        byteTrigger = OutputDateVarianceVola(varDateArray, _
                                                             varVarianceArray, _
                                                             varVolaArray, _
                                                             intRow_TradeDate, intColumn_TradeDate, _
                                                             intRow_VarianceDTE, intColumn_VarianceDTE, _
                                                             intRow_VolaDTE, intColumn_VolaDTE) 'Выводим.
                    Else
                        varVolaArray = varGetYangZhangJaggArray(varJoinMarketDataArray, _
                                                                byteFrequency, _
                                                                varLOT, _
                                                                intBegDateRowNum, _
                                                                intEndDateRowNum) 'Получаем массив массивов дат, дисп., вол..
                        byteTrigger = OutputComboDateVarianceVola(varVolaArray, _
                                                                  intRow_TradeDate, intColumn_TradeDate, _
                                                                  intRow_Variance100, intColumn_Variance100, _
                                                                  intRow_Vola100, intColumn_Vola100, _
                                                                  intRow_Variance50, intColumn_Variance50, _
                                                                  intRow_Vola50, intColumn_Vola50, _
                                                                  intRow_Variance20, intColumn_Variance20, _
                                                                  intRow_Vola20, intColumn_Vola20, _
                                                                  intRow_Variance10, intColumn_Variance10, _
                                                                  intRow_Vola10, intColumn_Vola10) 'Выводим.
                    End If
                End If
            End If
        End If
    End If
End Sub
