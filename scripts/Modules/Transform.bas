Attribute VB_Name = "Transform"
Option Explicit
Option Base 1

'Функция возвращает элементы в обратном порядке.
Private Function Reverse_(varArray As Variant) As Variant
    Dim varRevArray() As Variant
    Dim i As Integer, _
    intCount As Integer
    ReDim varRevArray(UBound(varArray))
    intCount = 0
    For i = UBound(varArray) To 1 Step -1
        intCount = intCount + 1
        varRevArray(intCount) = varArray(i, 1)
    Next i
    Reverse_ = varRevArray
End Function

'Функция выводит массив.
Private Function OutputArray(varArray As Variant)
    Dim i As Integer, _
    intRow As Integer, _
    intColumn As Integer
    Dim rngPoint As Range
    'Создаем ссылку на объект.
    Set rngPoint = ActiveCell
    'Получаем адрес объекта, т.е. точку отсчета для вывода.
    intRow = rngPoint.Row
    intColumn = rngPoint.Column
    'Уничтожаем объект.
    Set rngPoint = Nothing
    'Предварительная очистка ячеек, чтобы новые значения не накладывались на старые.
    ActiveSheet.Range(Cells(intRow, intColumn), Cells(intRow + (UBound(varArray) - 1), intColumn)).ClearContents
    'Вывод.
    For i = 1 To UBound(varArray)
        ActiveSheet.Cells(intRow + (i - 1), intColumn).Value = varArray(i)
    Next i
End Function

Public Sub Reverse()
    Attribute Reverse.VB_Description = "Поворачивает массив элементов на 180 C."
    Attribute Reverse.VB_ProcData.VB_Invoke_Func = "t\n14"
    Dim varArray As Variant, _
    varStart As Variant
    varArray = Application.InputBox("Укажите диапазон ячеек:", "Reversion: Запрос данных", "", Type:=8)
    varArray = Reverse_(varArray)
    varStart = OutputArray(varArray)
End Sub
