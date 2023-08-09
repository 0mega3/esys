Attribute VB_Name = "Weighting"
Option Explicit
Option Base 1

'Fishburne Weight Coefficients (FWC)
'Функция создает массив весов, используя формулу Фишберна (2*(n-i+1))/(n*(n+1)),
'где n - количество элементов в массиве, i - номер элемента в массиве.
Private Function FishbWeightCoeff(intQWE As Integer) As Variant
    Dim i As Integer
    Dim varFWC() As Variant, _
    varWeight As Variant
    ReDim varFWC(intQWE)
    For i = 1 To intQWE
        varWeight = (2 * (intQWE - i + 1)) / (intQWE * (intQWE + 1))
        varFWC(i) = varWeight
    Next i
    FishbWeightCoeff = varFWC
End Function

'Функция выводит массив весов.
Private Function OutputFishbWeightCoeff(varFWC As Variant)
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
    ActiveSheet.Range(Cells(intRow, intColumn), Cells(intRow + (UBound(varFWC) - 1), intColumn)).ClearContents
    'Вывод.
    For i = 1 To UBound(varFWC)
        ActiveSheet.Cells(intRow + (i - 1), intColumn).Value = varFWC(i)
    Next i
End Function

Public Sub FWC()
    Attribute FWC.VB_Description = "Создает массив весов, используя формулу Фишберна (2*(n-i+1))/(n*(n+1)), где n - кол-во элементов в массиве, i - номер элемента в массиве."
    Attribute FWC.VB_ProcData.VB_Invoke_Func = "w\n14"
    Dim intCount As Integer
    Dim varStartFWC As Variant, _
    varStartOutputFWC As Variant
    intCount = InputBox("Введите количество элементов:", "FishbWeightCoeff: Запрос данных", "")
    varStartFWC = FishbWeightCoeff(intCount)
    varStartOutputFWC = OutputFishbWeightCoeff(varStartFWC)
End Sub
