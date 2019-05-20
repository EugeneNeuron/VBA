Attribute VB_Name = "mCommon"
Option Explicit
'************************ Константы *****************************************
Public Const MSG_NO_DATA As String = "Соответствующие данные не найдены"
Public Const MSG_NO_VAR As String = "Существуют переменные. Измените значения переменных."
Public gstrTicket As String

Public Const gstrReport As String = "24 План ГКПЗ"  'Лист приложения
Public Const gstrBEx As String = "BEx"     'Лист bex запроса

Public Const gstrMapping As String = "mapping"  'Лист mapping
Public Const gstrReport_DP As String = "DP_ZSRM_O14_Q9_TMP"  'Daraprovider основных данных

Public Const gstrMappingArea As String = "MappingArea"  'Область мэппинга
Public Const REP_STA_ROW As Integer = 18
Public Const REP_STA_COL As Integer = 1
Public Const REP_END_COL As Integer = 22

'-------------Методы, ускоряющие обработку листов Excel----------
Public Sub Completion(ByRef vws As Worksheet)
    vws.Activate
    vws.Cells(1, 1).Select
    With ActiveWindow
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub

Public Sub Prepare()
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub
'-------------Методы, ускоряющие обработку листов Excel----------
'Проверка на наличие элемента в коллекции
Public Function InCollection(col As Collection, ByVal key As String) As Boolean
    Dim var As Variant
    Dim errNumber As Long

    InCollection = False
    Set var = Nothing

    Err.Clear
On Error Resume Next
    var = col.Item(key)
    errNumber = CLng(Err.Number)
On Error GoTo 0

    '5 is not in, 0 and 438 represent incollection
    If errNumber = 5 Then ' it is 5 if not in collection
        InCollection = False
    Else
        InCollection = True
    End If

End Function

'Очистка формата области
Public Sub ClearFormats(ByRef vRange As Range)
    If vRange Is Nothing Then Exit Sub
    With vRange
        'убираем границы
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        'убираем цвет
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
End Sub
'Установка границ по умолчанию у области
Public Sub ApplyBorders(ByRef vRange As Range)
    If vRange Is Nothing Then Exit Sub
    With vRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
End Sub
'Очистка области и перенос значение из другой области, если область-источник передается в метод
Public Sub ClearRangeContents(ByRef vRange As Range, _
                            Optional ByRef vSource As Range = Nothing, Optional ByVal vCopy As Boolean = False)
    If vRange Is Nothing Then Exit Sub
    vRange.ClearContents
    If vCopy = True Then
        If Not vSource Is Nothing Then
            vSource.Copy
            vRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            vRange.NumberFormat = "@"
        End If
    End If
End Sub
'Процедура копирования Range-а
Public Sub CopyRange(ByRef vToRange As Range, ByRef vFromRange As Range)
    vFromRange.Copy
    vToRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
End Sub






