VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "HCI3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Dim byteTrigger As Byte

Dim lmvarLastValue_HV100 As Variant, _
    lmvarLastValue_HV50 As Variant, _
    lmvarLastValue_HV20 As Variant, _
    lmvarLastValue_HV10 As Variant, _
    lmvarLastValue_IV As Variant

Dim lmvarArrayDeciles_HV100 As Variant, _
    lmvarArrayDeciles_HV50 As Variant, _
    lmvarArrayDeciles_HV20 As Variant, _
    lmvarArrayDeciles_HV10 As Variant, _
    lmvarArrayDeciles_IV As Variant

Private Function MainDecilesHV(rngArray_HV100 As Range, _
                               rngArray_HV50 As Range, _
                               rngArray_HV20 As Range, _
                               rngArray_HV10 As Range, _
                               intRow As Integer, _
                               intColumn As Integer) As Variant
    Dim varArray_HV100 As Variant, _
        varArray_HV50 As Variant, _
        varArray_HV20 As Variant, _
        varArray_HV10 As Variant

    Dim varMin_HV100 As Variant, _
        varMin_HV50 As Variant, _
        varMin_HV20 As Variant, _
        varMin_HV10 As Variant

    Dim varMax_HV100 As Variant, _
        varMax_HV50 As Variant, _
        varMax_HV20 As Variant, _
        varMax_HV10 As Variant

    Dim varQuantity_HV100 As Variant, _
        varQuantity_HV50 As Variant, _
        varQuantity_HV20 As Variant, _
        varQuantity_HV10 As Variant

    'Приводим диапазоны ячеек к двумерным массивам типа Variant.
    varArray_HV100 = rngArray_HV100
    varArray_HV50 = rngArray_HV50
    varArray_HV20 = rngArray_HV20
    varArray_HV10 = rngArray_HV10

    'Определяем последнее(текущее) значение в массивах HV.
    'Выносим на уровень модуля для дальнейшего сравнения с децилями
    'и расстановки цветовых меток.
    lmvarLastValue_HV100 = LastValue(varArray_HV100)
    lmvarLastValue_HV50 = LastValue(varArray_HV50)
    lmvarLastValue_HV20 = LastValue(varArray_HV20)
    lmvarLastValue_HV10 = LastValue(varArray_HV10)

    'Определяем количество значений не равных нулю в массивах HV.
    varQuantity_HV100 = Quantity(varArray_HV100)
    varQuantity_HV50 = Quantity(varArray_HV50)
    varQuantity_HV20 = Quantity(varArray_HV20)
    varQuantity_HV10 = Quantity(varArray_HV10)

    'Выбираем значения не равные нулю в массивах HV и сохраняем для последующей
    'сортировки.
    varArray_HV100 = Selection(varArray_HV100, varQuantity_HV100)
    varArray_HV50 = Selection(varArray_HV50, varQuantity_HV50)
    varArray_HV20 = Selection(varArray_HV20, varQuantity_HV20)
    varArray_HV10 = Selection(varArray_HV10, varQuantity_HV10)

    'Сортируем массивы HV по возрастанию (метод Пузырек).
    varArray_HV100 = SortBM(varArray_HV100)
    varArray_HV50 = SortBM(varArray_HV50)
    varArray_HV20 = SortBM(varArray_HV20)
    varArray_HV10 = SortBM(varArray_HV10)

    'Определяем значение максимума в массивах HV.
    varMax_HV100 = varArray_HV100(UBound(varArray_HV100))
    varMax_HV50 = varArray_HV50(UBound(varArray_HV50))
    varMax_HV20 = varArray_HV20(UBound(varArray_HV20))
    varMax_HV10 = varArray_HV10(UBound(varArray_HV10))

    'Определяем значение минимума в массивах HV.
    varMin_HV100 = varArray_HV100(1)
    varMin_HV50 = varArray_HV50(1)
    varMin_HV20 = varArray_HV20(1)
    varMin_HV10 = varArray_HV10(1)

    'Определяем массивы децилей HV.
    'Выносим на уровень модуля для дальнейшего вывода на экран и расстановки
    'цветовых меток.
    lmvarArrayDeciles_HV100 = СreatingDeciles(varQuantity_HV100, varArray_HV100, varMin_HV100, varMax_HV100)
    lmvarArrayDeciles_HV50 = СreatingDeciles(varQuantity_HV50, varArray_HV50, varMin_HV50, varMax_HV50)
    lmvarArrayDeciles_HV20 = СreatingDeciles(varQuantity_HV20, varArray_HV20, varMin_HV20, varMax_HV20)
    lmvarArrayDeciles_HV10 = СreatingDeciles(varQuantity_HV10, varArray_HV10, varMin_HV10, varMax_HV10)

    MainDecilesHV = OutputDecilesHV(lmvarArrayDeciles_HV100, _
                                    lmvarArrayDeciles_HV50, _
                                    lmvarArrayDeciles_HV20, _
                                    lmvarArrayDeciles_HV10, _
                                    intRow, intColumn)
End Function

Private Function MainDecilesIV(rngArray_IV As Range, _
                               intRow As Integer, _
                               intColumn As Integer) As Variant
    Dim varArray_IV As Variant, _
        varMin_IV As Variant, _
        varMax_IV As Variant, _
        intQuantity_IV As Integer

    'Приводим диапазон ячеек к двумерному массиву типа Variant.
    varArray_IV = rngArray_IV

    'Определяем последнее(текущее) значение в массиве IV.
    'Выносим на уровень модуля для дальнейшего сравнения с децилями
    'и расстановки цветовых меток.
    lmvarLastValue_IV = LastValue(varArray_IV)

    'Определяем количество значений не равных нулю в массиве IV.
    intQuantity_IV = Quantity(varArray_IV)

    'Выбираем значения не равные нулю в массиве IV и сохраняем для последующей
    'сортировки.
    varArray_IV = Selection(varArray_IV, intQuantity_IV)

    'Сортируем массив IV по возрастанию (метод Пузырек).
    varArray_IV = SortBM(varArray_IV)

    'Определяем значение максимума в массиве IV.
    varMax_IV = varArray_IV(UBound(varArray_IV))

    'Определяем значение минимума в массиве IV.
    varMin_IV = varArray_IV(1)

    'Определяем массив децилей IV.
    'Выносим на уровень модуля для дальнейшего вывода на экран и расстановки
    'цветовых меток.
    lmvarArrayDeciles_IV = СreatingDeciles(intQuantity_IV, varArray_IV, varMin_IV, varMax_IV)

    MainDecilesIV = OutputDecilesIV(lmvarArrayDeciles_IV, intRow, intColumn)
End Function

'Функция ищет последнее значение не равное нулю.
Private Function LastValue(varArrayVola As Variant) As Variant
    Dim i As Integer, intLV As Integer
    For i = 1 To UBound(varArrayVola)
        If IsNumeric(varArrayVola(i, 1)) = True And varArrayVola(i, 1) > 0 Then
            intLV = i
        End If
    Next i
    LastValue = varArrayVola(intLV, 1)
End Function

'Функция считает количество значений не равных нулю.
Private Function Quantity(varArrayVola As Variant) As Integer
    Dim i As Integer, intCount As Integer
    intCount = 0
    For i = 1 To UBound(varArrayVola)
        If IsNumeric(varArrayVola(i, 1)) = True And varArrayVola(i, 1) > 0 Then
            intCount = intCount + 1
        End If
    Next i
    Quantity = intCount
End Function

'Функция выбирает значения больше нуля и записывает в массив.
Private Function Selection(varArrayVola As Variant, varQuantityVola As Variant) As Variant
    Dim varSelectArrayVola As Variant
    Dim i As Integer, intCount As Integer
    ReDim varSelectArrayVola(varQuantityVola)
    intCount = 0
    For i = 1 To UBound(varArrayVola)
        If IsNumeric(varArrayVola(i, 1)) = True And varArrayVola(i, 1) > 0 Then
            intCount = intCount + 1
            varSelectArrayVola(intCount) = varArrayVola(i, 1)
        End If
    Next i
    Selection = varSelectArrayVola
End Function

'Функция сортирует массив по возрастанию (метод Пузырек).
Private Function SortBM(varArrayVola As Variant) As Variant
    Dim i As Integer
    Dim boolSwitch As Boolean
    Dim varTmp As Variant
    Do
        boolSwitch = 0
        For i = 1 To UBound(varArrayVola) - 1
            If varArrayVola(i + 1) < varArrayVola(i) Then
                boolSwitch = 1
                varTmp = varArrayVola(i)
                varArrayVola(i) = varArrayVola(i + 1)
                varArrayVola(i + 1) = varTmp
            End If
        Next i
        If boolSwitch = 0 Then Exit Do
    Loop
    SortBM = varArrayVola
End Function

'Функция вычисляет децили и записывает в массив.
'Если выражение i*n/10, где (i=1,2,3,4,5,6,7,8,9) не является целым числом,
'то i-я дециль является членом Xt ряда, где t самое меньшее из целых чисел,
'превышающих это выраженеие.
'1   2  3  4  5  6  7  8  9  10  11 - минимальное число членов ряда,
'если выражение i*n/10 не является целым числом.
'Min-D1-D2-D3-D4-D5-D6-D7-D8-D9-Max - расположение децилей.
'1   2  3  4  5  6  7  8  9  10  11 - количество элементов в массиве.
'Если же отношение i*n/10 - целое число, то i-я дециль может быть представлена
'каждым числом, содержащимся в интервале Xt и Xt+1, и, в частности,
'может быть равна средней арифметической этих двух членов ряда.
'n - количество членов ряда.
'1     2  3  4  5  6  7  8  9    10 - минимальное число членов ряда,
'если выражение i*n/10 является целым числом.
'Min-D1-D2-D3-D4-D5-D6-D7-D8-D9-Max - расположение децилей.
'1   2  3  4  5  6  7  8  9  10  11 - количество элементов в массиве.
Private Function СreatingDeciles(n As Variant, _
                                 varArrayVola As Variant, _
                                 varMinVola As Variant, _
                                 varMaxVola As Variant) As Variant
    Dim varArrayDecilesVola(11) As Variant
    Dim i As Integer
    For i = 1 To 9 'Диапазон равен количеству децилей, где (i=1,2,3,4,5,6,7,8,9).
        varArrayDecilesVola(i + 1) = i * n / 10 'Отсчет начинаем со 2 элемента по 10, чтобы зарезервировать 1 место за min, а 11 за max.
    Next i
    For i = 2 To 10
        If varArrayDecilesVola(i) <> Int(varArrayDecilesVola(i)) Then 'Если значение выражения i*n/10 есть дробное.
            varArrayDecilesVola(i) = Int(varArrayDecilesVola(i)) + 1
            varArrayDecilesVola(i) = varArrayVola(varArrayDecilesVola(i))
        Else 'Если значение выражения i*n/10 есть целое.
            varArrayDecilesVola(i) = (varArrayVola(varArrayDecilesVola(i)) + varArrayVola(varArrayDecilesVola(i) + 1)) / 2
        End If
    Next i
    varArrayDecilesVola(1) = varMinVola
    varArrayDecilesVola(11) = varMaxVola
    СreatingDeciles = varArrayDecilesVola
End Function

'Функция выводит массивы децилей HV.
'Min-D1-D2-D3-D4-D5-D6-D7-D8-D9-Max
Private Function OutputDecilesHV(varArrayDeciles_HV100 As Variant, _
                                 varArrayDeciles_HV50 As Variant, _
                                 varArrayDeciles_HV20 As Variant, _
                                 varArrayDeciles_HV10 As Variant, _
                                 intRow As Integer, _
                                 intColumn As Integer)
    Dim i As Integer

    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("VCone").Range(Cells(intRow + 4, intColumn + 1), Cells(intRow + 7, intColumn + 11)).ClearContents

    'Вывод.
    For i = 1 To 11
        Worksheets("VCone").Cells(intRow + 4, (intColumn + 1) + (i - 1)).Value = varArrayDeciles_HV10(i)
        Worksheets("VCone").Cells(intRow + 5, (intColumn + 1) + (i - 1)).Value = varArrayDeciles_HV20(i)
        Worksheets("VCone").Cells(intRow + 6, (intColumn + 1) + (i - 1)).Value = varArrayDeciles_HV50(i)
        Worksheets("VCone").Cells(intRow + 7, (intColumn + 1) + (i - 1)).Value = varArrayDeciles_HV100(i)
    Next i
End Function

'Функция выводит массив децилей IV.
'Min-D1-D2-D3-D4-D5-D6-D7-D8-D9-Max
Private Function OutputDecilesIV(varArrayDeciles_IV As Variant, _
                                 intRow As Integer, _
                                 intColumn As Integer)
    Dim i As Integer

    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    Worksheets("VCone").Range(Cells(intRow + 3, intColumn + 1), Cells(intRow + 3, intColumn + 11)).ClearContents

    'Вывод.
    For i = 1 To 11
        Worksheets("VCone").Cells(intRow + 3, (intColumn + 1) + (i - 1)).Value = varArrayDeciles_IV(i)
    Next i
End Function

'Функция расставляет цветовые метки на массивах децилей HV.
'Min-D1-D2-D3-D4-D5-D6-D7-D8-D9-Max
'Если последнее значение HV=Min, то Min закрашивается красным, чтобы было видно,
'что HV достигла нижнего предела.
'Если последнее значение HV=Max, то Max закрашивается красным, чтобы было видно,
'что HV достигла верхнего предела.
'Если min<HV<=D1,D1<HV<=D2,D2<HV<=D3,D3<HV<=D4,D4<HV<=D5,D5<HV<=D6,D6<HV<=D7,
'D7<HV<=D8,D8<HV<=D9,D9<HV<Max, то децили 1,2,3,4,5,6,7,8,9 и Max закрашиваются
'желтым цветом.
Private Function ColorDetectorDecilesHV(varArrayDeciles_HV100 As Variant, _
                                        varArrayDeciles_HV50 As Variant, _
                                        varArrayDeciles_HV20 As Variant, _
                                        varArrayDeciles_HV10 As Variant, _
                                        varLastValue_HV100 As Variant, _
                                        varLastValue_HV50 As Variant, _
                                        varLastValue_HV20 As Variant, _
                                        varLastValue_HV10 As Variant, _
                                        intRow As Integer, _
                                        intColumn As Integer)
    Dim i As Integer

    'Предварительное обесцвечивание ячеек, чтобы новые цветовые метки не накладывались на старые.
    Worksheets("VCone").Range(Cells(intRow + 4, intColumn + 1), Cells(intRow + 7, intColumn + 11)).Interior.ColorIndex = xlNone

    'Условия цветовых меток для HV.
    If varLastValue_HV10 <= varArrayDeciles_HV10(1) Then
        Worksheets("VCone").Cells(intRow + 4, intColumn + 1).Interior.ColorIndex = 3
    End If
    If varLastValue_HV10 >= varArrayDeciles_HV10(11) Then
        Worksheets("VCone").Cells(intRow + 4, intColumn + 11).Interior.ColorIndex = 3
    End If
    If varLastValue_HV10 > varArrayDeciles_HV10(10) And varLastValue_HV10 < varArrayDeciles_HV10(11) Then
        Worksheets("VCone").Cells(intRow + 4, intColumn + 11).Interior.ColorIndex = 6
    End If

    If varLastValue_HV20 <= varArrayDeciles_HV20(1) Then
        Worksheets("VCone").Cells(intRow + 5, intColumn + 1).Interior.ColorIndex = 3
    End If
    If varLastValue_HV20 >= varArrayDeciles_HV20(11) Then
        Worksheets("VCone").Cells(intRow + 5, intColumn + 11).Interior.ColorIndex = 3
    End If
    If varLastValue_HV20 > varArrayDeciles_HV20(10) And varLastValue_HV20 < varArrayDeciles_HV20(11) Then
        Worksheets("VCone").Cells(intRow + 5, intColumn + 11).Interior.ColorIndex = 6
    End If

    If varLastValue_HV50 <= varArrayDeciles_HV50(1) Then
        Worksheets("VCone").Cells(intRow + 6, intColumn + 1).Interior.ColorIndex = 3
    End If
    If varLastValue_HV50 >= varArrayDeciles_HV50(11) Then
        Worksheets("VCone").Cells(intRow + 6, intColumn + 11).Interior.ColorIndex = 3
    End If
    If varLastValue_HV50 > varArrayDeciles_HV50(10) And varLastValue_HV50 < varArrayDeciles_HV50(11) Then
        Worksheets("VCone").Cells(intRow + 6, intColumn + 11).Interior.ColorIndex = 6
    End If

    If varLastValue_HV100 <= varArrayDeciles_HV100(1) Then
        Worksheets("VCone").Cells(intRow + 7, intColumn + 1).Interior.ColorIndex = 3
    End If
    If varLastValue_HV100 >= varArrayDeciles_HV100(11) Then
        Worksheets("VCone").Cells(intRow + 7, intColumn + 11).Interior.ColorIndex = 3
    End If
    If varLastValue_HV100 > varArrayDeciles_HV100(10) And varLastValue_HV100 < varArrayDeciles_HV100(11) Then
        Worksheets("VCone").Cells(intRow + 7, intColumn + 11).Interior.ColorIndex = 6
    End If

    For i = 1 To 9
        If varLastValue_HV10 > varArrayDeciles_HV10(i) And varLastValue_HV10 <= varArrayDeciles_HV10(i + 1) Then
            Worksheets("VCone").Cells(intRow + 4, (intColumn + 1) + i).Interior.ColorIndex = 6
        End If
        If varLastValue_HV20 > varArrayDeciles_HV20(i) And varLastValue_HV20 <= varArrayDeciles_HV20(i + 1) Then
            Worksheets("VCone").Cells(intRow + 5, (intColumn + 1) + i).Interior.ColorIndex = 6
        End If
        If varLastValue_HV50 > varArrayDeciles_HV50(i) And varLastValue_HV50 <= varArrayDeciles_HV50(i + 1) Then
            Worksheets("VCone").Cells(intRow + 6, (intColumn + 1) + i).Interior.ColorIndex = 6
        End If
        If varLastValue_HV100 > varArrayDeciles_HV100(i) And varLastValue_HV100 <= varArrayDeciles_HV100(i + 1) Then
            Worksheets("VCone").Cells(intRow + 7, (intColumn + 1) + i).Interior.ColorIndex = 6
        End If
    Next i
End Function

'Функция расставляет цветовые метки на массиве децилей IV.
'Min-D1-D2-D3-D4-D5-D6-D7-D8-D9-Max
'Если последнее значение IV=Min, то Min закрашивается красным, чтобы было видно,
'что IV достигла нижнего предела.
'Если последнее значение IV=Max, то Max закрашивается красным, чтобы было видно,
'что IV достигла верхнего предела.
'Если min<IV<=D1,D1<IV<=D2,D2<IV<=D3,D3<IV<=D4,D4<IV<=D5,D5<IV<=D6,D6<IV<=D7,
'D7<IV<=D8,D8<IV<=D9,D9<IV<Max, то децили 1,2,3,4,5,6,7,8,9 и Max закрашиваются
'голубым цветом.
Private Function ColorDetectorDecilesIV(varArrayDeciles_IV As Variant, _
                                        varLastValue_IV As Variant, _
                                        intRow As Integer, _
                                        intColumn As Integer)
    Dim i As Integer

    'Предварительное обесцвечивание ячеек, чтобы новые цветовые метки не накладывались на старые.
    Worksheets("VCone").Range(Cells(intRow + 3, intColumn + 1), Cells(intRow + 3, intColumn + 11)).Interior.ColorIndex = xlNone

    'Условия цветовых меток для IV.
    If varLastValue_IV <= varArrayDeciles_IV(1) Then
        Worksheets("VCone").Cells(intRow + 3, intColumn + 1).Interior.ColorIndex = 3
    End If
    If varLastValue_IV >= varArrayDeciles_IV(11) Then
        Worksheets("VCone").Cells(intRow + 3, intColumn + 11).Interior.ColorIndex = 3
    End If
    If varLastValue_IV > varArrayDeciles_IV(10) And varLastValue_IV < varArrayDeciles_IV(11) Then
        Worksheets("VCone").Cells(intRow + 3, intColumn + 11).Interior.ColorIndex = 33
    End If

    For i = 1 To 9
        If varLastValue_IV > varArrayDeciles_IV(i) And varLastValue_IV <= varArrayDeciles_IV(i + 1) Then
            Worksheets("VCone").Cells(intRow + 3, (intColumn + 1) + i).Interior.ColorIndex = 33
        End If
    Next i
End Function

'Функция создает массив адресов строк, которые являются точками отсчета
'глубины охвата данных.
Private Function ReferencePointDepth(rngDateRange_Tool As Range) As Variant
    Dim colReferencePointDepth As Collection
    Set colReferencePointDepth = New Collection
    Set colReferencePointDepth = Nothing
    Dim varArrayReferencePointDepth As Variant

    Dim rngCell As Range
    Dim varValueCell As Variant
    Dim intYearСomparison As Integer
    Dim i As Integer

    For Each rngCell In rngDateRange_Tool
        varValueCell = rngCell.Value
        If Not colReferencePointDepth Is Nothing Then 'Если коллекция адресов строк не пуста.
            If IsEmpty(varValueCell) = False Then 'Если ячейка не пуста.
                If Year(varValueCell) > intYearСomparison Then 'Если наступил новый год.
                    colReferencePointDepth.Add (rngCell.Row) 'Добавляем адрес строки нового года.
                    intYearСomparison = Year(varValueCell) 'Перезаписываем год сравнения.
                End If
            Else
                colReferencePointDepth.Add (rngCell.Row - 1) 'Добавляем адрес строки последней даты текущего года т.е. определяем конечную точку массива дат.
                Exit For 'Прерываем значение цикла.
            End If
        ElseIf IsDate(varValueCell) Then 'Если коллекция адресов строк пуста и текущее значение есть дата.
            Set colReferencePointDepth = New Collection
            colReferencePointDepth.Add (rngCell.Row) 'Добавляем адрес строки первой даты истории дат т.е. определяем начальную точку массива дат.
            intYearСomparison = Year(varValueCell) 'Записываем год сравнения.
        End If
    Next
    ReDim varArrayReferencePointDepth(colReferencePointDepth.Count)
    For i = 1 To colReferencePointDepth.Count
        varArrayReferencePointDepth(i) = colReferencePointDepth(i)
    Next i
    ReferencePointDepth = varArrayReferencePointDepth
End Function

'Функция получает адрес конкретной ячеки с рабочего листа.
Private Function GetCellAddress(strWorksheet As String, _
                                strTool As String, _
                                intRow As Integer, _
                                intColumn As Integer)
    Dim i As Integer, j As Integer
    For i = 1 To 100
        For j = 1 To 100
            If IsError(Worksheets(strWorksheet).Cells(i, j).Value) = False Then
                If Worksheets(strWorksheet).Cells(i, j).Value = strTool Then
                    intRow = i
                    intColumn = j
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

'Функция назначает диапазоны HV.
Private Function RangeAsgmt_HV(strWorksheet As String, _
                               rngRange_HV100 As Range, _
                               rngRange_HV50 As Range, _
                               rngRange_HV20 As Range, _
                               rngRange_HV10 As Range, _
                               intRow1 As Integer, _
                               intRow2 As Integer, _
                               intColumn1 As Integer, _
                               intColumn2 As Integer, _
                               intColumn3 As Integer, _
                               intColumn4 As Integer)
    With Worksheets(strWorksheet)
        Set rngRange_HV100 = .Range(.Cells(intRow1, intColumn1), .Cells(intRow2, intColumn1))
        Set rngRange_HV50 = .Range(.Cells(intRow1, intColumn2), .Cells(intRow2, intColumn2))
        Set rngRange_HV20 = .Range(.Cells(intRow1, intColumn3), .Cells(intRow2, intColumn3))
        Set rngRange_HV10 = .Range(.Cells(intRow1, intColumn4), .Cells(intRow2, intColumn4))
    End With
End Function

'Функция назначает диапазоны IV.
Private Function RangeAsgmt_IV(strWorksheet As String, _
                               rngRange_IV As Range, _
                               intRow1 As Integer, _
                               intRow2 As Integer, _
                               intColumn1 As Integer)
    With Worksheets(strWorksheet)
        Set rngRange_IV = .Range(.Cells(intRow1, intColumn1), .Cells(intRow2, intColumn1))
    End With
End Function

Sub Decile_SBRF_Si_RI_BR_GD()
Attribute Decile_SBRF_Si_RI_BR_GD.VB_Description = "Выводит массивы волатильностей с разбивкой на децили и расставляет цветовые метки в тех децилях куда попадают последние текущие значения волатильностей.\n"
Attribute Decile_SBRF_Si_RI_BR_GD.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim intRow_SBRF As Integer, intColumn_SBRF As Integer, _
        intRow_Si As Integer, intColumn_Si As Integer, _
        intRow_RI As Integer, intColumn_RI As Integer, _
        intRow_BR As Integer, intColumn_BR As Integer, _
        intRow_GD As Integer, intColumn_GD As Integer

    Dim intRow_SBRF_MA_FrontIV_M As Integer, intColumn_SBRF_MA_FrontIV_M As Integer, _
        intRow_SBRF_MA_Combo2IV_M As Integer, intColumn_SBRF_MA_Combo2IV_M As Integer, _
        intRow_SBRF_MA_Combo3IV_M As Integer, intColumn_SBRF_MA_Combo3IV_M As Integer, _
        intRow_Si_MA_FrontIV_M As Integer, intColumn_Si_MA_FrontIV_M As Integer, _
        intRow_Si_MA_Combo2IV_M As Integer, intColumn_Si_MA_Combo2IV_M As Integer, _
        intRow_Si_MA_Combo3IV_M As Integer, intColumn_Si_MA_Combo3IV_M As Integer, _
        intRow_Si_MA_FrontIV_W As Integer, intColumn_Si_MA_FrontIV_W As Integer, _
        intRow_Si_MA_Combo2IV_W As Integer, intColumn_Si_MA_Combo2IV_W As Integer, _
        intRow_Si_MA_Combo3IV_W As Integer, intColumn_Si_MA_Combo3IV_W As Integer, _
        intRow_RI_MA_FrontIV_M As Integer, intColumn_RI_MA_FrontIV_M As Integer, _
        intRow_RI_MA_Combo2IV_M As Integer, intColumn_RI_MA_Combo2IV_M As Integer, _
        intRow_RI_MA_Combo3IV_M As Integer, intColumn_RI_MA_Combo3IV_M As Integer, _
        intRow_RI_MA_FrontIV_W As Integer, intColumn_RI_MA_FrontIV_W As Integer, _
        intRow_RI_MA_Combo2IV_W As Integer, intColumn_RI_MA_Combo2IV_W As Integer, _
        intRow_RI_MA_Combo3IV_W As Integer, intColumn_RI_MA_Combo3IV_W As Integer, _
        intRow_RVI_MA_Close As Integer, intColumn_RVI_MA_Close As Integer, _
        intRow_BR_MA_FrontIV_M As Integer, intColumn_BR_MA_FrontIV_M As Integer, _
        intRow_BR_MA_Combo2IV_M As Integer, intColumn_BR_MA_Combo2IV_M As Integer, _
        intRow_BR_MA_Combo3IV_M As Integer, intColumn_BR_MA_Combo3IV_M As Integer, _
        intRow_GD_MA_FrontIV_Q As Integer, intColumn_GD_MA_FrontIV_Q As Integer, _
        intRow_GD_MA_Combo2IV_Q As Integer, intColumn_GD_MA_Combo2IV_Q As Integer, _
        intRow_GD_MA_Combo3IV_Q As Integer, intColumn_GD_MA_Combo3IV_Q As Integer

    Dim varDepth_HV_SBRF As Variant, varDepth_IV_SBRF As Variant, varLifespan_SBRF As Variant, varCombo_SBRF As Variant, _
        varDepth_HV_Si As Variant, varDepth_IV_Si As Variant, varLifespan_Si As Variant, varCombo_Si As Variant, _
        varDepth_HV_RI As Variant, varDepth_IV_RI As Variant, varMode_IV_RI, varLifespan_RI As Variant, varCombo_RI As Variant, _
        varDepth_HV_BR As Variant, varDepth_IV_BR As Variant, varLifespan_BR As Variant, varCombo_BR As Variant, _
        varDepth_HV_GD As Variant, varDepth_IV_GD As Variant, varLifespan_GD As Variant, varCombo_GD As Variant

    Dim rngDateRange_HV_SBRF As Range, _
        rngDateRange_HV_Si As Range, _
        rngDateRange_HV_RI As Range, _
        rngDateRange_HV_BR As Range, _
        rngDateRange_HV_GD As Range, _
        rngDateRange_IV As Range

    Dim varReferencePointDepth_HV_SBRF As Variant, _
        varReferencePointDepth_HV_Si As Variant, _
        varReferencePointDepth_HV_RI As Variant, _
        varReferencePointDepth_HV_BR As Variant, _
        varReferencePointDepth_HV_GD As Variant, _
        varReferencePointDepth_IV As Variant

    Dim rngRange_HV100_SBRF As Range, rngRange_HV50_SBRF As Range, rngRange_HV20_SBRF As Range, rngRange_HV10_SBRF As Range, _
        rngRange_HV100_Si As Range, rngRange_HV50_Si As Range, rngRange_HV20_Si As Range, rngRange_HV10_Si As Range, _
        rngRange_HV100_RI As Range, rngRange_HV50_RI As Range, rngRange_HV20_RI As Range, rngRange_HV10_RI As Range, _
        rngRange_HV100_BR As Range, rngRange_HV50_BR As Range, rngRange_HV20_BR As Range, rngRange_HV10_BR As Range, _
        rngRange_HV100_GD As Range, rngRange_HV50_GD As Range, rngRange_HV20_GD As Range, rngRange_HV10_GD As Range

    Dim rngRange_IV_SBRF As Range, _
        rngRange_IV_Si As Range, _
        rngRange_IV_RI As Range, _
        rngRange_IV_BR As Range, _
        rngRange_IV_GD As Range

    Dim intDifferenceRPD As Integer, _
        intMaxDepth As Integer, _
        intRoundDepth As Integer, _
        intQuantityMembersRange As Integer

    Dim varRangetoArray As Variant

    'Находим адреса ячеек, т.е. находим точки отсчета для считывания значений глубин охвата данных и вывода массивов децилей, соответствующих этим глубинам.
    byteTrigger = GetCellAddress("VCone", "SBRF", intRow_SBRF, intColumn_SBRF)
    byteTrigger = GetCellAddress("VCone", "Si", intRow_Si, intColumn_Si)
    byteTrigger = GetCellAddress("VCone", "RI", intRow_RI, intColumn_RI)
    byteTrigger = GetCellAddress("VCone", "BR", intRow_BR, intColumn_BR)
    byteTrigger = GetCellAddress("VCone", "GD", intRow_GD, intColumn_GD)

    byteTrigger = GetCellAddress("IV", "SBRF_MA_FrontIV_M", intRow_SBRF_MA_FrontIV_M, intColumn_SBRF_MA_FrontIV_M)
    byteTrigger = GetCellAddress("IV", "SBRF_MA_Combo2IV_M", intRow_SBRF_MA_Combo2IV_M, intColumn_SBRF_MA_Combo2IV_M)
    byteTrigger = GetCellAddress("IV", "SBRF_MA_Combo3IV_M", intRow_SBRF_MA_Combo3IV_M, intColumn_SBRF_MA_Combo3IV_M)

    byteTrigger = GetCellAddress("IV", "Si_MA_FrontIV_M", intRow_Si_MA_FrontIV_M, intColumn_Si_MA_FrontIV_M)
    byteTrigger = GetCellAddress("IV", "Si_MA_Combo2IV_M", intRow_Si_MA_Combo2IV_M, intColumn_Si_MA_Combo2IV_M)
    byteTrigger = GetCellAddress("IV", "Si_MA_Combo3IV_M", intRow_Si_MA_Combo3IV_M, intColumn_Si_MA_Combo3IV_M)

    byteTrigger = GetCellAddress("IV", "Si_MA_FrontIV_W", intRow_Si_MA_FrontIV_W, intColumn_Si_MA_FrontIV_W)
    byteTrigger = GetCellAddress("IV", "Si_MA_Combo2IV_W", intRow_Si_MA_Combo2IV_W, intColumn_Si_MA_Combo2IV_W)
    byteTrigger = GetCellAddress("IV", "Si_MA_Combo3IV_W", intRow_Si_MA_Combo3IV_W, intColumn_Si_MA_Combo3IV_W)

    byteTrigger = GetCellAddress("IV", "RI_MA_FrontIV_M", intRow_RI_MA_FrontIV_M, intColumn_RI_MA_FrontIV_M)
    byteTrigger = GetCellAddress("IV", "RI_MA_Combo2IV_M", intRow_RI_MA_Combo2IV_M, intColumn_RI_MA_Combo2IV_M)
    byteTrigger = GetCellAddress("IV", "RI_MA_Combo3IV_M", intRow_RI_MA_Combo3IV_M, intColumn_RI_MA_Combo3IV_M)

    byteTrigger = GetCellAddress("IV", "RI_MA_FrontIV_W", intRow_RI_MA_FrontIV_W, intColumn_RI_MA_FrontIV_W)
    byteTrigger = GetCellAddress("IV", "RI_MA_Combo2IV_W", intRow_RI_MA_Combo2IV_W, intColumn_RI_MA_Combo2IV_W)
    byteTrigger = GetCellAddress("IV", "RI_MA_Combo3IV_W", intRow_RI_MA_Combo3IV_W, intColumn_RI_MA_Combo3IV_W)

    byteTrigger = GetCellAddress("IV", "RVI_MA_Close", intRow_RVI_MA_Close, intColumn_RVI_MA_Close)

    byteTrigger = GetCellAddress("IV", "BR_MA_FrontIV_M", intRow_BR_MA_FrontIV_M, intColumn_BR_MA_FrontIV_M)
    byteTrigger = GetCellAddress("IV", "BR_MA_Combo2IV_M", intRow_BR_MA_Combo2IV_M, intColumn_BR_MA_Combo2IV_M)
    byteTrigger = GetCellAddress("IV", "BR_MA_Combo3IV_M", intRow_BR_MA_Combo3IV_M, intColumn_BR_MA_Combo3IV_M)

    byteTrigger = GetCellAddress("IV", "GD_MA_FrontIV_Q", intRow_GD_MA_FrontIV_Q, intColumn_GD_MA_FrontIV_Q)
    byteTrigger = GetCellAddress("IV", "GD_MA_Combo2IV_Q", intRow_GD_MA_Combo2IV_Q, intColumn_GD_MA_Combo2IV_Q)
    byteTrigger = GetCellAddress("IV", "GD_MA_Combo3IV_Q", intRow_GD_MA_Combo3IV_Q, intColumn_GD_MA_Combo3IV_Q)

    'Считываем значения с панели пресетов.
    varDepth_HV_SBRF = Worksheets("VCone").Cells(intRow_SBRF, intColumn_SBRF + 5).Value
    varDepth_HV_Si = Worksheets("VCone").Cells(intRow_Si, intColumn_Si + 5).Value
    varDepth_HV_RI = Worksheets("VCone").Cells(intRow_RI, intColumn_RI + 3).Value
    varDepth_HV_BR = Worksheets("VCone").Cells(intRow_BR, intColumn_BR + 5).Value
    varDepth_HV_GD = Worksheets("VCone").Cells(intRow_GD, intColumn_GD + 5).Value

    varDepth_IV_SBRF = Worksheets("VCone").Cells(intRow_SBRF, intColumn_SBRF + 7).Value
    varDepth_IV_Si = Worksheets("VCone").Cells(intRow_Si, intColumn_Si + 7).Value
    varDepth_IV_RI = Worksheets("VCone").Cells(intRow_RI, intColumn_RI + 5).Value
    varDepth_IV_BR = Worksheets("VCone").Cells(intRow_BR, intColumn_BR + 7).Value
    varDepth_IV_GD = Worksheets("VCone").Cells(intRow_GD, intColumn_GD + 7).Value

    varMode_IV_RI = Worksheets("VCone").Cells(intRow_RI, intColumn_RI + 7).Value

    varLifespan_SBRF = Worksheets("VCone").Cells(intRow_SBRF, intColumn_SBRF + 9).Value
    varLifespan_Si = Worksheets("VCone").Cells(intRow_Si, intColumn_Si + 9).Value
    varLifespan_RI = Worksheets("VCone").Cells(intRow_RI, intColumn_RI + 9).Value
    varLifespan_BR = Worksheets("VCone").Cells(intRow_BR, intColumn_BR + 9).Value
    varLifespan_GD = Worksheets("VCone").Cells(intRow_GD, intColumn_GD + 9).Value

    varCombo_SBRF = Worksheets("VCone").Cells(intRow_SBRF, intColumn_SBRF + 11).Value
    varCombo_Si = Worksheets("VCone").Cells(intRow_Si, intColumn_Si + 11).Value
    varCombo_RI = Worksheets("VCone").Cells(intRow_RI, intColumn_RI + 11).Value
    varCombo_BR = Worksheets("VCone").Cells(intRow_BR, intColumn_BR + 11).Value
    varCombo_GD = Worksheets("VCone").Cells(intRow_GD, intColumn_GD + 11).Value

    'Назначаем диапазоны дат.
    Set rngDateRange_HV_SBRF = Worksheets("SBRF_HV").Range("A1:A20000")
    Set rngDateRange_HV_Si = Worksheets("Si_HV").Range("A1:A20000")
    Set rngDateRange_HV_RI = Worksheets("RI_HV").Range("A1:A20000")
    Set rngDateRange_HV_BR = Worksheets("BR_HV").Range("A1:A20000")
    Set rngDateRange_HV_GD = Worksheets("GD_HV").Range("A1:A20000")

    Set rngDateRange_IV = Worksheets("IV").Range("A1:A20000")

    'SBRF
    '1.Если параметры Depth_HV и Depth_IV не имеют значений.
    If IsEmpty(varDepth_HV_SBRF) = True And IsEmpty(varDepth_IV_SBRF) = True Then
        'Очищаем и обесцвечиваем область IV и HV.
        Worksheets("VCone").Range(Cells(intRow_SBRF + 3, intColumn_SBRF + 1), Cells(intRow_SBRF + 7, intColumn_SBRF + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_SBRF + 3, intColumn_SBRF + 1), Cells(intRow_SBRF + 7, intColumn_SBRF + 11)).Interior.ColorIndex = xlNone
        '2.Если параметр Depth_HV имеет значение, а параметр Depth_IV нет.
    ElseIf IsEmpty(varDepth_HV_SBRF) = False And IsEmpty(varDepth_IV_SBRF) = True Then
        'Очищаем и обесцвечиваем область IV.
        Worksheets("VCone").Range(Cells(intRow_SBRF + 3, intColumn_SBRF + 1), Cells(intRow_SBRF + 3, intColumn_SBRF + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_SBRF + 3, intColumn_SBRF + 1), Cells(intRow_SBRF + 3, intColumn_SBRF + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_HV_SBRF = ReferencePointDepth(rngDateRange_HV_SBRF) 'Создаем контрольные точки глубины(КТГ или RPD) доступа к данным HV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_SBRF) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_SBRF) = True And varDepth_HV_SBRF > 0 And varDepth_HV_SBRF <= intMaxDepth Or LCase(varDepth_HV_SBRF) = "max" Then
            If IsNumeric(varDepth_HV_SBRF) = True Then 'Если значение глубины есть число.
                If varDepth_HV_SBRF = Int(varDepth_HV_SBRF) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("SBRF_HV", _
                                                rngRange_HV100_SBRF, _
                                                rngRange_HV50_SBRF, _
                                                rngRange_HV20_SBRF, _
                                                rngRange_HV10_SBRF, _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF) - varDepth_HV_SBRF)), _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF)) - varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("SBRF_HV", _
                                                rngRange_HV100_SBRF, _
                                                rngRange_HV50_SBRF, _
                                                rngRange_HV20_SBRF, _
                                                rngRange_HV10_SBRF, _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("SBRF_HV", _
                                            rngRange_HV100_SBRF, _
                                            rngRange_HV50_SBRF, _
                                            rngRange_HV20_SBRF, _
                                            rngRange_HV10_SBRF, _
                                            CInt(varReferencePointDepth_HV_SBRF(1)), _
                                            CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_SBRF 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_SBRF, _
                                            rngRange_HV50_SBRF, _
                                            rngRange_HV20_SBRF, _
                                            rngRange_HV10_SBRF, _
                                            intRow_SBRF, intColumn_SBRF)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_SBRF, intColumn_SBRF)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "SBRF" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "SBRF" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '3.Если параметр Depth_HV не имеет значение, а параметр Depth_IV да.
    ElseIf IsEmpty(varDepth_HV_SBRF) = True And IsEmpty(varDepth_IV_SBRF) = False Then
        'Очищаем и обесцвечиваем область HV.
        Worksheets("VCone").Range(Cells(intRow_SBRF + 4, intColumn_SBRF + 1), Cells(intRow_SBRF + 7, intColumn_SBRF + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_SBRF + 4, intColumn_SBRF + 1), Cells(intRow_SBRF + 7, intColumn_SBRF + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_SBRF) = True And varDepth_IV_SBRF > 0 And varDepth_IV_SBRF <= intMaxDepth Or LCase(varDepth_IV_SBRF) = "max" Then
            If LCase(varLifespan_SBRF) = "m" Then 'Если опционы месячные (M).
                '-------------------------------------------------------------------------------------
                If varCombo_SBRF = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_SBRF) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_SBRF = Int(varDepth_IV_SBRF) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_SBRF)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_FrontIV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_SBRF, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_SBRF_MA_FrontIV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_SBRF = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_SBRF) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_SBRF = Int(varDepth_IV_SBRF) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_SBRF)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_SBRF, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_SBRF_MA_Combo2IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_SBRF = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_SBRF) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_SBRF = Int(varDepth_IV_SBRF) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_SBRF)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_SBRF, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_SBRF_MA_Combo3IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_SBRF) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "SBRF" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "SBRF" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_SBRF) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "SBRF" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "SBRF" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            End If

            If Not rngRange_IV_SBRF Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_SBRF 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_SBRF, intRow_SBRF, intColumn_SBRF)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_SBRF, intColumn_SBRF)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "SBRF" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "SBRF" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '4.Если параметры Depth_HV и Depth_IV имеют значения.
    ElseIf IsEmpty(varDepth_HV_SBRF) = False And IsEmpty(varDepth_IV_SBRF) = False Then
        varReferencePointDepth_HV_SBRF = ReferencePointDepth(rngDateRange_HV_SBRF) 'Создаем КТГ доступа к данным HV.
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_SBRF) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_SBRF) = True And varDepth_HV_SBRF > 0 And varDepth_HV_SBRF <= intMaxDepth Or LCase(varDepth_HV_SBRF) = "max" Then
            If IsNumeric(varDepth_HV_SBRF) = True Then 'Если значение глубины есть число.
                If varDepth_HV_SBRF = Int(varDepth_HV_SBRF) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("SBRF_HV", _
                                                rngRange_HV100_SBRF, _
                                                rngRange_HV50_SBRF, _
                                                rngRange_HV20_SBRF, _
                                                rngRange_HV10_SBRF, _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF) - varDepth_HV_SBRF)), _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF)) - varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("SBRF_HV", _
                                                rngRange_HV100_SBRF, _
                                                rngRange_HV50_SBRF, _
                                                rngRange_HV20_SBRF, _
                                                rngRange_HV10_SBRF, _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("SBRF_HV", _
                                            rngRange_HV100_SBRF, _
                                            rngRange_HV50_SBRF, _
                                            rngRange_HV20_SBRF, _
                                            rngRange_HV10_SBRF, _
                                            CInt(varReferencePointDepth_HV_SBRF(1)), _
                                            CInt(varReferencePointDepth_HV_SBRF(UBound(varReferencePointDepth_HV_SBRF))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_SBRF 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_SBRF, _
                                            rngRange_HV50_SBRF, _
                                            rngRange_HV20_SBRF, _
                                            rngRange_HV10_SBRF, _
                                            intRow_SBRF, intColumn_SBRF)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_SBRF, intColumn_SBRF)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "SBRF" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "SBRF" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_SBRF) = True And varDepth_IV_SBRF > 0 And varDepth_IV_SBRF <= intMaxDepth Or LCase(varDepth_IV_SBRF) = "max" Then
            If LCase(varLifespan_SBRF) = "m" Then 'Если опционы месячные (M).
                '-------------------------------------------------------------------------------------
                If varCombo_SBRF = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_SBRF) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_SBRF = Int(varDepth_IV_SBRF) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_SBRF)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_FrontIV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_SBRF, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_SBRF_MA_FrontIV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_SBRF = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_SBRF) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_SBRF = Int(varDepth_IV_SBRF) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_SBRF)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_SBRF, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_SBRF_MA_Combo2IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_SBRF = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_SBRF) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_SBRF = Int(varDepth_IV_SBRF) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_SBRF)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_SBRF) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_SBRF / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_SBRF, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_SBRF_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_SBRF, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_SBRF_MA_Combo3IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_SBRF) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "SBRF" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "SBRF" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_SBRF) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "SBRF" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "SBRF" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            End If

            If Not rngRange_IV_SBRF Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_SBRF 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_SBRF, intRow_SBRF, intColumn_SBRF)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_SBRF, intColumn_SBRF)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "SBRF" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "SBRF" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
    End If

    'Si
    '1.Если параметры Depth_HV и Depth_IV не имеют значений.
    If IsEmpty(varDepth_HV_Si) = True And IsEmpty(varDepth_IV_Si) = True Then
        'Очищаем и обесцвечиваем область IV и HV.
        Worksheets("VCone").Range(Cells(intRow_Si + 3, intColumn_Si + 1), Cells(intRow_Si + 7, intColumn_Si + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_Si + 3, intColumn_Si + 1), Cells(intRow_Si + 7, intColumn_Si + 11)).Interior.ColorIndex = xlNone
        '2.Если параметр Depth_HV имеет значение, а параметр Depth_IV нет.
    ElseIf IsEmpty(varDepth_HV_Si) = False And IsEmpty(varDepth_IV_Si) = True Then
        'Очищаем и обесцвечиваем область IV.
        Worksheets("VCone").Range(Cells(intRow_Si + 3, intColumn_Si + 1), Cells(intRow_Si + 3, intColumn_Si + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_Si + 3, intColumn_Si + 1), Cells(intRow_Si + 3, intColumn_Si + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_HV_Si = ReferencePointDepth(rngDateRange_HV_Si) 'Создаем контрольные точки глубины(КТГ или RPD) доступа к данным HV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_Si) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_Si) = True And varDepth_HV_Si > 0 And varDepth_HV_Si <= intMaxDepth Or LCase(varDepth_HV_Si) = "max" Then
            If IsNumeric(varDepth_HV_Si) = True Then 'Если значение глубины есть число.
                If varDepth_HV_Si = Int(varDepth_HV_Si) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("Si_HV", _
                                                rngRange_HV100_Si, _
                                                rngRange_HV50_Si, _
                                                rngRange_HV20_Si, _
                                                rngRange_HV10_Si, _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si) - varDepth_HV_Si)), _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si))), _
                                                15, 20, 25, 30) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si)) - varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("Si_HV", _
                                                rngRange_HV100_Si, _
                                                rngRange_HV50_Si, _
                                                rngRange_HV20_Si, _
                                                rngRange_HV10_Si, _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si))), _
                                                15, 20, 25, 30) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("Si_HV", _
                                            rngRange_HV100_Si, _
                                            rngRange_HV50_Si, _
                                            rngRange_HV20_Si, _
                                            rngRange_HV10_Si, _
                                            CInt(varReferencePointDepth_HV_Si(1)), _
                                            CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si))), _
                                            15, 20, 25, 30) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_Si 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_Si, _
                                            rngRange_HV50_Si, _
                                            rngRange_HV20_Si, _
                                            rngRange_HV10_Si, _
                                            intRow_Si, intColumn_Si)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_Si, intColumn_Si)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Si" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "Si" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '3.Если параметр Depth_HV не имеет значение, а параметр Depth_IV да.
    ElseIf IsEmpty(varDepth_HV_Si) = True And IsEmpty(varDepth_IV_Si) = False Then
        'Очищаем и обесцвечиваем область HV.
        Worksheets("VCone").Range(Cells(intRow_Si + 4, intColumn_Si + 1), Cells(intRow_Si + 7, intColumn_Si + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_Si + 4, intColumn_Si + 1), Cells(intRow_Si + 7, intColumn_Si + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_Si) = True And varDepth_IV_Si > 0 And varDepth_IV_Si <= intMaxDepth Or LCase(varDepth_IV_Si) = "max" Then
            If LCase(varLifespan_Si) = "m" Then 'Если опционы месячные (M).
                '-------------------------------------------------------------------------------------
                If varCombo_Si = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_FrontIV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo2IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo3IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_Si) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf LCase(varLifespan_Si) = "w" Then 'Если опционы недельные (W).
                '-------------------------------------------------------------------------------------
                If varCombo_Si = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_W) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_W) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_FrontIV_W) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_W) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_W) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo2IV_W) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_W) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_W) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo3IV_W) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_Si) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_Si) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Si" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Si" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
            End If

            If Not rngRange_IV_Si Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_Si 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_Si, intRow_Si, intColumn_Si)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_Si, intColumn_Si)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "Si" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '4.Если параметры Depth_HV и Depth_IV имеют значения.
    ElseIf IsEmpty(varDepth_HV_Si) = False And IsEmpty(varDepth_IV_Si) = False Then
        varReferencePointDepth_HV_Si = ReferencePointDepth(rngDateRange_HV_Si) 'Создаем КТГ доступа к данным HV.
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_Si) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_Si) = True And varDepth_HV_Si > 0 And varDepth_HV_Si <= intMaxDepth Or LCase(varDepth_HV_Si) = "max" Then
            If IsNumeric(varDepth_HV_Si) = True Then 'Если значение глубины есть число.
                If varDepth_HV_Si = Int(varDepth_HV_Si) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("Si_HV", _
                                                rngRange_HV100_Si, _
                                                rngRange_HV50_Si, _
                                                rngRange_HV20_Si, _
                                                rngRange_HV10_Si, _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si) - varDepth_HV_Si)), _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si))), _
                                                15, 20, 25, 30) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si)) - varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("Si_HV", _
                                                rngRange_HV100_Si, _
                                                rngRange_HV50_Si, _
                                                rngRange_HV20_Si, _
                                                rngRange_HV10_Si, _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si))), _
                                                15, 20, 25, 30) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("Si_HV", _
                                            rngRange_HV100_Si, _
                                            rngRange_HV50_Si, _
                                            rngRange_HV20_Si, _
                                            rngRange_HV10_Si, _
                                            CInt(varReferencePointDepth_HV_Si(1)), _
                                            CInt(varReferencePointDepth_HV_Si(UBound(varReferencePointDepth_HV_Si))), _
                                            15, 20, 25, 30) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_Si 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_Si, _
                                            rngRange_HV50_Si, _
                                            rngRange_HV20_Si, _
                                            rngRange_HV10_Si, _
                                            intRow_Si, intColumn_Si)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_Si, intColumn_Si)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Si" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "Si" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_Si) = True And varDepth_IV_Si > 0 And varDepth_IV_Si <= intMaxDepth Or LCase(varDepth_IV_Si) = "max" Then
            If LCase(varLifespan_Si) = "m" Then 'Если опционы месячные (M).
                '-------------------------------------------------------------------------------------
                If varCombo_Si = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_FrontIV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo2IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo3IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_Si) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf LCase(varLifespan_Si) = "w" Then 'Если опционы недельные (W).
                '-------------------------------------------------------------------------------------
                If varCombo_Si = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_W) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_FrontIV_W) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_FrontIV_W) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_W) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo2IV_W) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo2IV_W) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_Si = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_Si) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_Si = Int(varDepth_IV_Si) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_Si)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_W) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_Si) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_Si / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_Si, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_Si_MA_Combo3IV_W) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_Si, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_Si_MA_Combo3IV_W) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_Si) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_Si) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Si" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "Si" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
            End If

            If Not rngRange_IV_Si Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_Si 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_Si, intRow_Si, intColumn_Si)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_Si, intColumn_Si)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "Si" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "Si" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
    End If

    'RI
    '1.Если параметры Depth_HV и Depth_IV не имеют значений.
    If IsEmpty(varDepth_HV_RI) = True And IsEmpty(varDepth_IV_RI) = True Then
        'Очищаем и обесцвечиваем область IV и HV.
        Worksheets("VCone").Range(Cells(intRow_RI + 3, intColumn_RI + 1), Cells(intRow_RI + 7, intColumn_RI + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_RI + 3, intColumn_RI + 1), Cells(intRow_RI + 7, intColumn_RI + 11)).Interior.ColorIndex = xlNone
        '2.Если параметр Depth_HV имеет значение, а параметр Depth_IV нет.
    ElseIf IsEmpty(varDepth_HV_RI) = False And IsEmpty(varDepth_IV_RI) = True Then
        'Очищаем и обесцвечиваем область IV.
        Worksheets("VCone").Range(Cells(intRow_RI + 3, intColumn_RI + 1), Cells(intRow_RI + 3, intColumn_RI + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_RI + 3, intColumn_RI + 1), Cells(intRow_RI + 3, intColumn_RI + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_HV_RI = ReferencePointDepth(rngDateRange_HV_RI) 'Создаем контрольные точки глубины(КТГ или RPD) доступа к данным HV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_RI) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_RI) = True And varDepth_HV_RI > 0 And varDepth_HV_RI <= intMaxDepth Or LCase(varDepth_HV_RI) = "max" Then
            If IsNumeric(varDepth_HV_RI) = True Then 'Если значение глубины есть число.
                If varDepth_HV_RI = Int(varDepth_HV_RI) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("RI_HV", _
                                                rngRange_HV100_RI, _
                                                rngRange_HV50_RI, _
                                                rngRange_HV20_RI, _
                                                rngRange_HV10_RI, _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI) - varDepth_HV_RI)), _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI)) - varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("RI_HV", _
                                                rngRange_HV100_RI, _
                                                rngRange_HV50_RI, _
                                                rngRange_HV20_RI, _
                                                rngRange_HV10_RI, _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("RI_HV", _
                                            rngRange_HV100_RI, _
                                            rngRange_HV50_RI, _
                                            rngRange_HV20_RI, _
                                            rngRange_HV10_RI, _
                                            CInt(varReferencePointDepth_HV_RI(1)), _
                                            CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_RI 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_RI, _
                                            rngRange_HV50_RI, _
                                            rngRange_HV20_RI, _
                                            rngRange_HV10_RI, _
                                            intRow_RI, intColumn_RI)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_RI, intColumn_RI)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "RI" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "RI" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '3.Если параметр Depth_HV не имеет значение, а параметр Depth_IV да.
    ElseIf IsEmpty(varDepth_HV_RI) = True And IsEmpty(varDepth_IV_RI) = False Then
        'Очищаем и обесцвечиваем область HV.
        Worksheets("VCone").Range(Cells(intRow_RI + 4, intColumn_RI + 1), Cells(intRow_RI + 7, intColumn_RI + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_RI + 4, intColumn_RI + 1), Cells(intRow_RI + 7, intColumn_RI + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_RI) = True And varDepth_IV_RI > 0 And varDepth_IV_RI <= intMaxDepth Or LCase(varDepth_IV_RI) = "max" Then
            If LCase(varMode_IV_RI) = "iv" Then
                If LCase(varLifespan_RI) = "m" Then 'Если опционы месячные (M).
                    '-------------------------------------------------------------------------------------
                    If varCombo_RI = 1 Then 'Если нет склеивания (Front).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_M) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_M) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 2 Then 'Если склеены две ближайших серии (Front+Next).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_M) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_M) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_M) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_M) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf IsEmpty(varCombo_RI) = True Then 'Если параметр Combo пуст.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите значение параметра Combo." _
                        & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    Else 'Если параметр Combo некорректен.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите корректное значение параметра Combo." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    End If
                ElseIf LCase(varLifespan_RI) = "w" Then 'Если опционы недельные (W).
                    '-------------------------------------------------------------------------------------
                    If varCombo_RI = 1 Then 'Если нет склеивания (Front).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_W) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_W) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_FrontIV_W) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 2 Then 'Если склеены две ближайших серии (Front+Next).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_W) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_W) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo2IV_W) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_W) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_W) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo3IV_W) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf IsEmpty(varCombo_RI) = True Then 'Если параметр Combo пуст.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите значение параметра Combo." _
                        & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    Else 'Если параметр Combo некорректен.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите корректное значение параметра Combo." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    End If
                ElseIf IsEmpty(varLifespan_RI) = True Then 'Если параметр Lifespan пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "RI" _
                    & vbNewLine & "Введите значение параметра Lifespan." _
                    & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                    & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
                Else 'Если параметр Lifespan некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "RI" _
                    & vbNewLine & "Введите корректное значение параметра Lifespan." _
                    & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
                End If
            ElseIf LCase(varMode_IV_RI) = "rvi" Then
                '=================================
                If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                    If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_RI, _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_RVI_MA_Close) 'Назначаем диапазон IV.
                    Else 'Если значение глубины есть дробное.
                        intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                        'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                        intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                        intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_RI, _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_RVI_MA_Close) 'Назначаем диапазон IV.
                    End If
                Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                    byteTrigger = RangeAsgmt_IV("IV", _
                                                rngRange_IV_RI, _
                                                CInt(varReferencePointDepth_IV(1)), _
                                                CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                intColumn_RVI_MA_Close) 'Назначаем диапазон IV.
                End If
            ElseIf IsEmpty(varMode_IV_RI) = True Then 'Если параметр Mode_IV пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "RI" _
                & vbNewLine & "Введите значение параметра Mode_IV." _
                & vbNewLine & "Параметр Mode_IV определяет режим рыночной волатильности." _
                & vbNewLine & "Доступные значения параметра Mode_IV - IV, RVI.")
            Else 'Если параметр Mode_IV некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "RI" _
                & vbNewLine & "Введите корректное значение параметра Mode_IV." _
                & vbNewLine & "Доступные значения параметра Mode_IV - IV, RVI.")
            End If

            If Not rngRange_IV_RI Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_RI 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_RI, intRow_RI, intColumn_RI)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_RI, intColumn_RI)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "RI" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "RI" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '4.Если параметры Depth_HV и Depth_IV имеют значения.
    ElseIf IsEmpty(varDepth_HV_RI) = False And IsEmpty(varDepth_IV_RI) = False Then
        varReferencePointDepth_HV_RI = ReferencePointDepth(rngDateRange_HV_RI) 'Создаем КТГ доступа к данным HV.
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_RI) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_RI) = True And varDepth_HV_RI > 0 And varDepth_HV_RI <= intMaxDepth Or LCase(varDepth_HV_RI) = "max" Then
            If IsNumeric(varDepth_HV_RI) = True Then 'Если значение глубины есть число.
                If varDepth_HV_RI = Int(varDepth_HV_RI) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("RI_HV", _
                                                rngRange_HV100_RI, _
                                                rngRange_HV50_RI, _
                                                rngRange_HV20_RI, _
                                                rngRange_HV10_RI, _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI) - varDepth_HV_RI)), _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI)) - varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("RI_HV", _
                                                rngRange_HV100_RI, _
                                                rngRange_HV50_RI, _
                                                rngRange_HV20_RI, _
                                                rngRange_HV10_RI, _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("RI_HV", _
                                            rngRange_HV100_RI, _
                                            rngRange_HV50_RI, _
                                            rngRange_HV20_RI, _
                                            rngRange_HV10_RI, _
                                            CInt(varReferencePointDepth_HV_RI(1)), _
                                            CInt(varReferencePointDepth_HV_RI(UBound(varReferencePointDepth_HV_RI))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_RI 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_RI, _
                                            rngRange_HV50_RI, _
                                            rngRange_HV20_RI, _
                                            rngRange_HV10_RI, _
                                            intRow_RI, intColumn_RI)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_RI, intColumn_RI)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "RI" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "RI" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_RI) = True And varDepth_IV_RI > 0 And varDepth_IV_RI <= intMaxDepth Or LCase(varDepth_IV_RI) = "max" Then
            If LCase(varMode_IV_RI) = "iv" Then
                If LCase(varLifespan_RI) = "m" Then 'Если опционы месячные (M).
                    '-------------------------------------------------------------------------------------
                    If varCombo_RI = 1 Then 'Если нет склеивания (Front).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_M) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_M) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 2 Then 'Если склеены две ближайших серии (Front+Next).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_M) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_M) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_M) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_M) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf IsEmpty(varCombo_RI) = True Then 'Если параметр Combo пуст.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите значение параметра Combo." _
                        & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    Else 'Если параметр Combo некорректен.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите корректное значение параметра Combo." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    End If
                ElseIf LCase(varLifespan_RI) = "w" Then 'Если опционы недельные (W).
                    '-------------------------------------------------------------------------------------
                    If varCombo_RI = 1 Then 'Если нет склеивания (Front).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_W) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_FrontIV_W) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_FrontIV_W) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 2 Then 'Если склеены две ближайших серии (Front+Next).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_W) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo2IV_W) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo2IV_W) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf varCombo_RI = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                        If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                            If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_W) 'Назначаем диапазон IV.
                            Else 'Если значение глубины есть дробное.
                                intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                                'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                                intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                                intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                                byteTrigger = RangeAsgmt_IV("IV", _
                                                            rngRange_IV_RI, _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                            CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                            intColumn_RI_MA_Combo3IV_W) 'Назначаем диапазон IV.
                            End If
                        Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_RI, _
                                                        CInt(varReferencePointDepth_IV(1)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_RI_MA_Combo3IV_W) 'Назначаем диапазон IV.
                        End If
                        '-------------------------------------------------------------------------------------
                    ElseIf IsEmpty(varCombo_RI) = True Then 'Если параметр Combo пуст.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите значение параметра Combo." _
                        & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    Else 'Если параметр Combo некорректен.
                        MsgBox ("Оба-на!!!" _
                        & vbNewLine & "RI" _
                        & vbNewLine & "Введите корректное значение параметра Combo." _
                        & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                    End If
                ElseIf IsEmpty(varLifespan_RI) = True Then 'Если параметр Lifespan пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "RI" _
                    & vbNewLine & "Введите значение параметра Lifespan." _
                    & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                    & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
                Else 'Если параметр Lifespan некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "RI" _
                    & vbNewLine & "Введите корректное значение параметра Lifespan." _
                    & vbNewLine & "Доступные значения параметра Lifespan - M (месячные), W (недельные).")
                End If
            ElseIf LCase(varMode_IV_RI) = "rvi" Then
                If IsNumeric(varDepth_IV_RI) = True Then 'Если значение глубины есть число.
                    If varDepth_IV_RI = Int(varDepth_IV_RI) Then 'Если значение глубины есть целое.
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_RI, _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_RI)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_RVI_MA_Close) 'Назначаем диапазон IV.
                    Else 'Если значение глубины есть дробное.
                        intRoundDepth = Int(varDepth_IV_RI) + 1 'Увеличиваем значение глубины до ближайшего целого.
                        'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                        intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                        intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_RI / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_RI, _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_RVI_MA_Close) 'Назначаем диапазон IV.
                    End If
                Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                    byteTrigger = RangeAsgmt_IV("IV", _
                                                rngRange_IV_RI, _
                                                CInt(varReferencePointDepth_IV(1)), _
                                                CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                intColumn_RVI_MA_Close) 'Назначаем диапазон IV.
                End If
            ElseIf IsEmpty(varMode_IV_RI) = True Then 'Если параметр Mode_IV пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "RI" _
                & vbNewLine & "Введите значение параметра Mode_IV." _
                & vbNewLine & "Параметр Mode_IV определяет режим рыночной волатильности." _
                & vbNewLine & "Доступные значения параметра Mode_IV - IV, RVI.")
            Else 'Если параметр Mode_IV некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "RI" _
                & vbNewLine & "Введите корректное значение параметра Mode_IV." _
                & vbNewLine & "Доступные значения параметра Mode_IV - IV, RVI.")
            End If

            If Not rngRange_IV_RI Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_RI 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_RI, intRow_RI, intColumn_RI)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_RI, intColumn_RI)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "RI" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "RI" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
    End If

    'BR
    '1.Если параметры Depth_HV и Depth_IV не имеют значений.
    If IsEmpty(varDepth_HV_BR) = True And IsEmpty(varDepth_IV_BR) = True Then
        'Очищаем и обесцвечиваем область IV и HV.
        Worksheets("VCone").Range(Cells(intRow_BR + 3, intColumn_BR + 1), Cells(intRow_BR + 7, intColumn_BR + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_BR + 3, intColumn_BR + 1), Cells(intRow_BR + 7, intColumn_BR + 11)).Interior.ColorIndex = xlNone
        '2.Если параметр Depth_HV имеет значение, а параметр Depth_IV нет.
    ElseIf IsEmpty(varDepth_HV_BR) = False And IsEmpty(varDepth_IV_BR) = True Then
        'Очищаем и обесцвечиваем область IV.
        Worksheets("VCone").Range(Cells(intRow_BR + 3, intColumn_BR + 1), Cells(intRow_BR + 3, intColumn_BR + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_BR + 3, intColumn_BR + 1), Cells(intRow_BR + 3, intColumn_BR + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_HV_BR = ReferencePointDepth(rngDateRange_HV_BR) 'Создаем контрольные точки глубины(КТГ или RPD) доступа к данным HV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_BR) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_BR) = True And varDepth_HV_BR > 0 And varDepth_HV_BR <= intMaxDepth Or LCase(varDepth_HV_BR) = "max" Then
            If IsNumeric(varDepth_HV_BR) = True Then 'Если значение глубины есть число.
                If varDepth_HV_BR = Int(varDepth_HV_BR) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("BR_HV", _
                                                rngRange_HV100_BR, _
                                                rngRange_HV50_BR, _
                                                rngRange_HV20_BR, _
                                                rngRange_HV10_BR, _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR) - varDepth_HV_BR)), _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR)) - varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("BR_HV", _
                                                rngRange_HV100_BR, _
                                                rngRange_HV50_BR, _
                                                rngRange_HV20_BR, _
                                                rngRange_HV10_BR, _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("BR_HV", _
                                            rngRange_HV100_BR, _
                                            rngRange_HV50_BR, _
                                            rngRange_HV20_BR, _
                                            rngRange_HV10_BR, _
                                            CInt(varReferencePointDepth_HV_BR(1)), _
                                            CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_BR 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_BR, _
                                            rngRange_HV50_BR, _
                                            rngRange_HV20_BR, _
                                            rngRange_HV10_BR, _
                                            intRow_BR, intColumn_BR)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_BR, intColumn_BR)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "BR" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "BR" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '3.Если параметр Depth_HV не имеет значение, а параметр Depth_IV да.
    ElseIf IsEmpty(varDepth_HV_BR) = True And IsEmpty(varDepth_IV_BR) = False Then
        'Очищаем и обесцвечиваем область HV.
        Worksheets("VCone").Range(Cells(intRow_BR + 4, intColumn_BR + 1), Cells(intRow_BR + 7, intColumn_BR + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_BR + 4, intColumn_BR + 1), Cells(intRow_BR + 7, intColumn_BR + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_BR) = True And varDepth_IV_BR > 0 And varDepth_IV_BR <= intMaxDepth Or LCase(varDepth_IV_BR) = "max" Then
            If LCase(varLifespan_BR) = "m" Then 'Если опционы месячные (M).
                '-------------------------------------------------------------------------------------
                If varCombo_BR = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_BR) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_BR = Int(varDepth_IV_BR) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_BR)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_FrontIV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_BR, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_BR_MA_FrontIV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_BR = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_BR) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_BR = Int(varDepth_IV_BR) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_BR)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_BR, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_BR_MA_Combo2IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_BR = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_BR) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_BR = Int(varDepth_IV_BR) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_BR)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_BR, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_BR_MA_Combo3IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_BR) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "BR" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "BR" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_BR) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "BR" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "BR" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            End If

            If Not rngRange_IV_BR Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_BR 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_BR, intRow_BR, intColumn_BR)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_BR, intColumn_BR)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "BR" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "BR" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '4.Если параметры Depth_HV и Depth_IV имеют значения.
    ElseIf IsEmpty(varDepth_HV_BR) = False And IsEmpty(varDepth_IV_BR) = False Then
        varReferencePointDepth_HV_BR = ReferencePointDepth(rngDateRange_HV_BR) 'Создаем КТГ доступа к данным HV.
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_BR) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_BR) = True And varDepth_HV_BR > 0 And varDepth_HV_BR <= intMaxDepth Or LCase(varDepth_HV_BR) = "max" Then
            If IsNumeric(varDepth_HV_BR) = True Then 'Если значение глубины есть число.
                If varDepth_HV_BR = Int(varDepth_HV_BR) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("BR_HV", _
                                                rngRange_HV100_BR, _
                                                rngRange_HV50_BR, _
                                                rngRange_HV20_BR, _
                                                rngRange_HV10_BR, _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR) - varDepth_HV_BR)), _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR)) - varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("BR_HV", _
                                                rngRange_HV100_BR, _
                                                rngRange_HV50_BR, _
                                                rngRange_HV20_BR, _
                                                rngRange_HV10_BR, _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("BR_HV", _
                                            rngRange_HV100_BR, _
                                            rngRange_HV50_BR, _
                                            rngRange_HV20_BR, _
                                            rngRange_HV10_BR, _
                                            CInt(varReferencePointDepth_HV_BR(1)), _
                                            CInt(varReferencePointDepth_HV_BR(UBound(varReferencePointDepth_HV_BR))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_BR 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_BR, _
                                            rngRange_HV50_BR, _
                                            rngRange_HV20_BR, _
                                            rngRange_HV10_BR, _
                                            intRow_BR, intColumn_BR)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_BR, intColumn_BR)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "BR" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "BR" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_BR) = True And varDepth_IV_BR > 0 And varDepth_IV_BR <= intMaxDepth Or LCase(varDepth_IV_BR) = "max" Then
            If LCase(varLifespan_BR) = "m" Then 'Если опционы месячные (M).
                '-------------------------------------------------------------------------------------
                If varCombo_BR = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_BR) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_BR = Int(varDepth_IV_BR) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_BR)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_FrontIV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_FrontIV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_BR, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_BR_MA_FrontIV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_BR = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_BR) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_BR = Int(varDepth_IV_BR) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_BR)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo2IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_BR, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_BR_MA_Combo2IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_BR = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_BR) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_BR = Int(varDepth_IV_BR) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_BR)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_BR) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_BR / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_BR, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_BR_MA_Combo3IV_M) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_BR, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_BR_MA_Combo3IV_M) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_BR) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "BR" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "BR" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_BR) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "BR" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "BR" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - M (месячные).")
            End If

            If Not rngRange_IV_BR Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_BR 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_BR, intRow_BR, intColumn_BR)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_BR, intColumn_BR)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "BR" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "BR" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
    End If

    'GD
    '1.Если параметры Depth_HV и Depth_IV не имеют значений.
    If IsEmpty(varDepth_HV_GD) = True And IsEmpty(varDepth_IV_GD) = True Then
        'Очищаем и обесцвечиваем область IV и HV.
        Worksheets("VCone").Range(Cells(intRow_GD + 3, intColumn_GD + 1), Cells(intRow_GD + 7, intColumn_GD + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_GD + 3, intColumn_GD + 1), Cells(intRow_GD + 7, intColumn_GD + 11)).Interior.ColorIndex = xlNone
        '2.Если параметр Depth_HV имеет значение, а параметр Depth_IV нет.
    ElseIf IsEmpty(varDepth_HV_GD) = False And IsEmpty(varDepth_IV_GD) = True Then
        'Очищаем и обесцвечиваем область IV.
        Worksheets("VCone").Range(Cells(intRow_GD + 3, intColumn_GD + 1), Cells(intRow_GD + 3, intColumn_GD + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_GD + 3, intColumn_GD + 1), Cells(intRow_GD + 3, intColumn_GD + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_HV_GD = ReferencePointDepth(rngDateRange_HV_GD) 'Создаем контрольные точки глубины(КТГ или RPD) доступа к данным HV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_GD) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_GD) = True And varDepth_HV_GD > 0 And varDepth_HV_GD <= intMaxDepth Or LCase(varDepth_HV_GD) = "max" Then
            If IsNumeric(varDepth_HV_GD) = True Then 'Если значение глубины есть число.
                If varDepth_HV_GD = Int(varDepth_HV_GD) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("GD_HV", _
                                                rngRange_HV100_GD, _
                                                rngRange_HV50_GD, _
                                                rngRange_HV20_GD, _
                                                rngRange_HV10_GD, _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD) - varDepth_HV_GD)), _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD)) - varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("GD_HV", _
                                                rngRange_HV100_GD, _
                                                rngRange_HV50_GD, _
                                                rngRange_HV20_GD, _
                                                rngRange_HV10_GD, _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("GD_HV", _
                                            rngRange_HV100_GD, _
                                            rngRange_HV50_GD, _
                                            rngRange_HV20_GD, _
                                            rngRange_HV10_GD, _
                                            CInt(varReferencePointDepth_HV_GD(1)), _
                                            CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_GD 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_GD, _
                                            rngRange_HV50_GD, _
                                            rngRange_HV20_GD, _
                                            rngRange_HV10_GD, _
                                            intRow_GD, intColumn_GD)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_GD, intColumn_GD)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "GD" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "GD" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '3.Если параметр Depth_HV не имеет значение, а параметр Depth_IV да.
    ElseIf IsEmpty(varDepth_HV_GD) = True And IsEmpty(varDepth_IV_GD) = False Then
        'Очищаем и обесцвечиваем область HV.
        Worksheets("VCone").Range(Cells(intRow_GD + 4, intColumn_GD + 1), Cells(intRow_GD + 7, intColumn_GD + 11)).ClearContents
        Worksheets("VCone").Range(Cells(intRow_GD + 4, intColumn_GD + 1), Cells(intRow_GD + 7, intColumn_GD + 11)).Interior.ColorIndex = xlNone
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_GD) = True And varDepth_IV_GD > 0 And varDepth_IV_GD <= intMaxDepth Or LCase(varDepth_IV_GD) = "max" Then
            If LCase(varLifespan_GD) = "q" Then 'Если опционы квартальные (Q).
                '-------------------------------------------------------------------------------------
                If varCombo_GD = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_GD) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_GD = Int(varDepth_IV_GD) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_GD)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_FrontIV_Q) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_FrontIV_Q) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_GD, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_GD_MA_FrontIV_Q) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_GD = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_GD) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_GD = Int(varDepth_IV_GD) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_GD)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo2IV_Q) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo2IV_Q) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_GD, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_GD_MA_Combo2IV_Q) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_GD = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_GD) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_GD = Int(varDepth_IV_GD) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_GD)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo3IV_Q) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo3IV_Q) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_GD, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_GD_MA_Combo3IV_Q) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_GD) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "GD" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "GD" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_GD) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "GD" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - Q (квартальные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "GD" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - Q (квартальные).")
            End If

            If Not rngRange_IV_GD Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_GD 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_GD, intRow_GD, intColumn_GD)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_GD, intColumn_GD)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "GD" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "GD" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        '4.Если параметры Depth_HV и Depth_IV имеют значения.
    ElseIf IsEmpty(varDepth_HV_GD) = False And IsEmpty(varDepth_IV_GD) = False Then
        varReferencePointDepth_HV_GD = ReferencePointDepth(rngDateRange_HV_GD) 'Создаем КТГ доступа к данным HV.
        varReferencePointDepth_IV = ReferencePointDepth(rngDateRange_IV) 'Создаем КТГ доступа к данным IV.
        'HV
        intMaxDepth = UBound(varReferencePointDepth_HV_GD) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_HV_GD) = True And varDepth_HV_GD > 0 And varDepth_HV_GD <= intMaxDepth Or LCase(varDepth_HV_GD) = "max" Then
            If IsNumeric(varDepth_HV_GD) = True Then 'Если значение глубины есть число.
                If varDepth_HV_GD = Int(varDepth_HV_GD) Then 'Если значение глубины есть целое.
                    byteTrigger = RangeAsgmt_HV("GD_HV", _
                                                rngRange_HV100_GD, _
                                                rngRange_HV50_GD, _
                                                rngRange_HV20_GD, _
                                                rngRange_HV10_GD, _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD) - varDepth_HV_GD)), _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                Else 'Если значение глубины есть дробное.
                    intRoundDepth = Int(varDepth_HV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                    'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                    intDifferenceRPD = varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD)) - varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD) - intRoundDepth)
                    intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_HV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                    byteTrigger = RangeAsgmt_HV("GD_HV", _
                                                rngRange_HV100_GD, _
                                                rngRange_HV50_GD, _
                                                rngRange_HV20_GD, _
                                                rngRange_HV10_GD, _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD)) - intDifferenceRPD), _
                                                CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD))), _
                                                7, 12, 17, 22) 'Назначаем диапазоны HV.
                End If
            Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                byteTrigger = RangeAsgmt_HV("GD_HV", _
                                            rngRange_HV100_GD, _
                                            rngRange_HV50_GD, _
                                            rngRange_HV20_GD, _
                                            rngRange_HV10_GD, _
                                            CInt(varReferencePointDepth_HV_GD(1)), _
                                            CInt(varReferencePointDepth_HV_GD(UBound(varReferencePointDepth_HV_GD))), _
                                            7, 12, 17, 22) 'Назначаем диапазоны HV.
            End If

            'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
            varRangetoArray = rngRange_HV100_GD 'Приводим диапазон к двумерному массиву типа Variant.
            intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
            If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                'Выводим массивы децилей HV.
                byteTrigger = MainDecilesHV(rngRange_HV100_GD, _
                                            rngRange_HV50_GD, _
                                            rngRange_HV20_GD, _
                                            rngRange_HV10_GD, _
                                            intRow_GD, intColumn_GD)
                'Расставляем цветовые метки массивов децилей HV.
                byteTrigger = ColorDetectorDecilesHV(lmvarArrayDeciles_HV100, lmvarArrayDeciles_HV50, lmvarArrayDeciles_HV20, lmvarArrayDeciles_HV10, _
                                                     lmvarLastValue_HV100, lmvarLastValue_HV50, lmvarLastValue_HV20, lmvarLastValue_HV10, _
                                                     intRow_GD, intColumn_GD)
            Else
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "GD" _
                & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                & vbNewLine & "Нужно ввести большее значение для параметра Depth_HV." _
                & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "GD" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_HV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_HV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
        'IV
        intMaxDepth = UBound(varReferencePointDepth_IV) - 1 'Замеряем максимум глубины охвата данных за минусом 1 КТГ.
        'Если глубина есть число в диапазоне 0<глубина<=максимальная глубина или текст равный префиксу max(максимальное значение глубины).
        If IsNumeric(varDepth_IV_GD) = True And varDepth_IV_GD > 0 And varDepth_IV_GD <= intMaxDepth Or LCase(varDepth_IV_GD) = "max" Then
            If LCase(varLifespan_GD) = "q" Then 'Если опционы квартальные (Q).
                '-------------------------------------------------------------------------------------
                If varCombo_GD = 1 Then 'Если нет склеивания (Front).
                    If IsNumeric(varDepth_IV_GD) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_GD = Int(varDepth_IV_GD) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_GD)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_FrontIV_Q) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_FrontIV_Q) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_GD, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_GD_MA_FrontIV_Q) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_GD = 2 Then 'Если склеены две ближайших серии (Front+Next).
                    If IsNumeric(varDepth_IV_GD) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_GD = Int(varDepth_IV_GD) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_GD)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo2IV_Q) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo2IV_Q) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_GD, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_GD_MA_Combo2IV_Q) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf varCombo_GD = 3 Then 'Если склеены три ближайших серии (Front+Next+Back).
                    If IsNumeric(varDepth_IV_GD) = True Then 'Если значение глубины есть число.
                        If varDepth_IV_GD = Int(varDepth_IV_GD) Then 'Если значение глубины есть целое.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - varDepth_IV_GD)), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo3IV_Q) 'Назначаем диапазон IV.
                        Else 'Если значение глубины есть дробное.
                            intRoundDepth = Int(varDepth_IV_GD) + 1 'Увеличиваем значение глубины до ближайшего целого.
                            'Находим разницу между конечной КТГ и этой же КТГ сдвинутой на величину intRoundDepth.
                            intDifferenceRPD = varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - varReferencePointDepth_IV(UBound(varReferencePointDepth_IV) - intRoundDepth)
                            intDifferenceRPD = FormatNumber(intDifferenceRPD * (varDepth_IV_GD / intRoundDepth), 0) 'Приводим целое intDifferenceRPD к дробному т.е. находим промежуточную точку КТГ.
                            byteTrigger = RangeAsgmt_IV("IV", _
                                                        rngRange_IV_GD, _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV)) - intDifferenceRPD), _
                                                        CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                        intColumn_GD_MA_Combo3IV_Q) 'Назначаем диапазон IV.
                        End If
                    Else 'Если значение глубины есть текст равный префиксу max(максимальное значение глубины).
                        byteTrigger = RangeAsgmt_IV("IV", _
                                                    rngRange_IV_GD, _
                                                    CInt(varReferencePointDepth_IV(1)), _
                                                    CInt(varReferencePointDepth_IV(UBound(varReferencePointDepth_IV))), _
                                                    intColumn_GD_MA_Combo3IV_Q) 'Назначаем диапазон IV.
                    End If
                    '-------------------------------------------------------------------------------------
                ElseIf IsEmpty(varCombo_GD) = True Then 'Если параметр Combo пуст.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "GD" _
                    & vbNewLine & "Введите значение параметра Combo." _
                    & vbNewLine & "Параметр Combo определяет кол-во ближайших серий опционов для склеивания." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                Else 'Если параметр Combo некорректен.
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "GD" _
                    & vbNewLine & "Введите корректное значение параметра Combo." _
                    & vbNewLine & "Доступные значения параметра Combo - 1 (Front), 2 (Next), 3 (Back).")
                End If
            ElseIf IsEmpty(varLifespan_GD) = True Then 'Если параметр Lifespan пуст.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "GD" _
                & vbNewLine & "Введите значение параметра Lifespan." _
                & vbNewLine & "Параметр Lifespan определяет срок обращения опционов." _
                & vbNewLine & "Доступные значения параметра Lifespan - Q (квартальные).")
            Else 'Если параметр Lifespan некорректен.
                MsgBox ("Оба-на!!!" _
                & vbNewLine & "GD" _
                & vbNewLine & "Введите корректное значение параметра Lifespan." _
                & vbNewLine & "Доступные значения параметра Lifespan - Q (квартальные).")
            End If

            If Not rngRange_IV_GD Is Nothing Then 'Если диапазон IV назначен.
                'Проверяем достаточность количества членов ряда при расчете децилей (минимальное условие достаточности=10).
                varRangetoArray = rngRange_IV_GD 'Приводим диапазон к двумерному массиву типа Variant.
                intQuantityMembersRange = Quantity(varRangetoArray) 'Считаем количество членов ряда.
                If intQuantityMembersRange >= 10 Then 'Если условие достаточности соблюдается.
                    'Выводим массив децилей IV.
                    byteTrigger = MainDecilesIV(rngRange_IV_GD, intRow_GD, intColumn_GD)
                    'Расставляем цветовые метки массива децилей IV.
                    byteTrigger = ColorDetectorDecilesIV(lmvarArrayDeciles_IV, lmvarLastValue_IV, intRow_GD, intColumn_GD)
                Else
                    MsgBox ("Оба-на!!!" _
                    & vbNewLine & "GD" _
                    & vbNewLine & "Недостаточно данных для расчета значений децилей." _
                    & vbNewLine & "Нужно ввести большее значение для параметра Depth_IV." _
                    & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
                End If
            End If
        Else
            MsgBox ("Оба-на!!!" _
            & vbNewLine & "GD" _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV<0." _
            & vbNewLine & "Возможно, текущее значение параметра Depth_IV>максимума." _
            & vbNewLine & "Максимальное значение параметра Depth_IV=" & intMaxDepth & "." _
            & vbNewLine & "Ввод префикса max будет = максимальной глубине охвата данных.")
        End If
    End If
End Sub
