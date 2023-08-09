VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Option Compare Text
Dim byteTrigger As Byte

Private Sub Workbook_Open()
    byteTrigger = ExpDateValidator()
End Sub

Private Function ExpDateValidator()
    Dim intRow_ExpDate As Integer, _
        intColumn_ExpDate As Integer, _
        strExpDate As String, _
        intDiffDTE As Integer
    'The cell address of the ExpDate field is calculated.
    byteTrigger = GetAdressParam(intRow_ExpDate, intColumn_ExpDate, "ExpDate")
    'Reading the value of the ExpDate field
    strExpDate = Worksheets("AutoHV").Cells(intRow_ExpDate, intColumn_ExpDate + 1).Value
    If strExpDate = "TBU" Then
    Else
        intDiffDTE = DateDiff("d", Now, CDate(strExpDate)) + 1
        If intDiffDTE <= 0 Then
            With Worksheets("AutoHV").Cells(intRow_ExpDate, intColumn_ExpDate + 1)
                .Interior.Color = RGB(237, 72, 48) 'Santa Claus Hat
                .Value = "TBU" ' To Be Updated
            End With
        End If
    End If
End Function

'The cell address of the input parameter is calculated.
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
