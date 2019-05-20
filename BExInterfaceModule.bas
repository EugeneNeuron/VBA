Attribute VB_Name = "BExInterfaceModule"
Option Explicit
'Описываем глобальные для модуля константы и переменные
Public gintGridCntr As Integer '– счетчик обновленных Grid-ов
Public gcolDP As New Collection '– коллекция с параметрами датапровайдеров

Public Const LOT_ID As Integer = 3
Public Const POS_ID As Integer = 4

Public Function BEx() As Object
  Set BEx = Application.Run("BExAnalyzer.xla!GetBEx", ThisWorkbook)
End Function

Public Sub CallBack(ParamArray varname())
    Dim GC As Integer
    Dim lngRow As Long, lngPrevLastRow As Long

    Dim objRprtWS As Worksheet, objBExWS As Worksheet
    Dim rngReport As Range, rngTemp As Range

    'Dim arrMapping() As Variant
    Dim objDP As DPItem
    Dim objGKPZ As clGKPZ
    
    Dim colO14 As New Collection
    
    Dim strTmp As String

On Error GoTo stop_macros
    'При первом вызове макроса инициализируем коллекцию и подсчитываем в статической функции количество Grid-ов в рабочей книге, как указывалось выше, именно столько раз будет вызван макрос.
    ' обновление первого провайдера данных
    If IsMissing(BEx) Then
        Exit Sub
    End If
    
    If (gintGridCntr = 0) Then
        Call Prepare
        Set gcolDP = Nothing
        Set gcolDP = New Collection
        GC = GridCount(True, BEx)
    End If

    'Элементы в коллекции должны иметь уникальное имя. В этом качестве мы будем использовать имя датапровайдера. Производим проверку на наличие элемента в коллекции с помощью функции ExistInCollection (т.к. в коллекции стандартно не предусмотрен метод Exist, то мы написали свою булевскую функцию проверки существования элемента в коллекции). Если элемента с проверяемым именем в коллекции нет, то создаем новый с типом ссылающимся на наш класс DPItem, таким образом элемент унаследует все свойства класса.
    If Not (InCollection(gcolDP, varname(0))) Then
        gcolDP.Add New DPItem, varname(0)
        'Заполняем свойства текущего элемента значениями.
        gcolDP.Item(varname(0)).DPName = varname(0)
        gcolDP.Item(varname(0)).isEmpty = True
        'Заполнение остальных свойств мы вынесли в отдельную процедуру FillgcolDP.
        ' заполнение оставшихся атрибутов текущего провайдера данных в коллекции
        Call FillgcolDP(varname(0), varname(1), varname(2))
        gcolDP.Item(varname(0)).isUpdated = True
    End If

    'Увеличиваем счетчик обновленных Grid-ов
    gintGridCntr = gintGridCntr + 1
    If (gintGridCntr = GridCount) Then
       'Подготовка данных к следующему обновлению рабочей книги
        gintGridCntr = 0

        Set objRprtWS = Worksheets(gstrReport)
        Set objBExWS = wsO14
        Set objDP = gcolDP(gstrReport_DP)

        'Если данных нет, не обновляем версионные данные.
        If objBExWS.Cells(objDP.DPOffsetY, objDP.DPOffsetX) = MSG_NO_VAR Or _
           objBExWS.Cells(objDP.DPOffsetY, objDP.DPOffsetX) = MSG_NO_DATA Then
            objDP.dataRowsCount = 1
        End If

        lngPrevLastRow = objRprtWS.Cells.SpecialCells(xlLastCell).Row
        If lngPrevLastRow > REP_STA_ROW Then
            'Очищаем область ниже полученных данных(область предыдущего построения)
            Set rngTemp = objRprtWS.Range(objRprtWS.Cells(REP_STA_ROW, REP_STA_COL), _
                                            objRprtWS.Cells(lngPrevLastRow, REP_END_COL))
            rngTemp.Delete Shift:=xlUp
            Set rngTemp = Nothing
        End If
        
        If objDP.LastRow >= objDP.FirstRow And objBExWS.Cells(objDP.DPOffsetY, objDP.DPOffsetX) <> MSG_NO_VAR And _
           objBExWS.Cells(objDP.DPOffsetY, objDP.DPOffsetX) <> MSG_NO_DATA Then
           
            Dim RR As Long ' строки в исходники
           
            Dim sKey As String
            
            Dim dblTemp As Double
            'Проходим все строки
            For lngRow = objDP.FirstRow To objDP.LastRow
               
               
                sKey = CStr(wsO14.Cells(lngRow, 4))
                If InCollection(colO14, sKey) = False Then
                
                    Set objGKPZ = New clGKPZ
                    objGKPZ.sPosition = wsO14.Cells(lngRow, 4)
                    objGKPZ.sLotID = wsO14.Cells(lngRow, 3)
                    objGKPZ.sTitleName = CaptionCon(wsO14, lngRow, 20, 8)
                    objGKPZ.sNetPower = wsO14.Cells(lngRow, 19)
                    objGKPZ.sName = CaptionCon(wsO14, lngRow, 12, 5)
                    objGKPZ.sOrgBuy = wsO14.Cells(lngRow, 7)
                    objGKPZ.sOrgBuy2 = wsO14.Cells(lngRow, 2)
                    objGKPZ.sPlanPositionSum = wsO14.Cells(lngRow, 30)
                    objGKPZ.sAttachType = wsO14.Cells(lngRow, 5)
                    objGKPZ.sMethodBuy = wsO14.Cells(lngRow, 6)
                    objGKPZ.sDateStartProcess = wsO14.Cells(lngRow, 28)
                    objGKPZ.sDateCloseADeal = wsO14.Cells(lngRow, 11)
                    objGKPZ.sPlanDate = wsO14.Cells(lngRow, 8)
                    objGKPZ.sRowPos = colO14.Count + 1

                    colO14.Add objGKPZ, sKey
                    Set objGKPZ = Nothing
                End If
            Next lngRow
            
            'Заполнение инв. проектов у заявки/позиции START
            For lngRow = 2 To ZR1DS405.Range("A" & Rows.Count).End(xlUp).Row
                strTmp = CStr(ZR1DS405.Cells(lngRow, 1))
                If InCollection(colO14, strTmp) = True And CStr(ZR1DS405.Cells(lngRow, 2)) <> "#" Then
                    If colO14(strTmp).sProjCode = "" Then
                        colO14(strTmp).sProjCode = CStr(ZR1DS405.Cells(lngRow, 2))
                    ElseIf InStr(1, CStr(colO14(strTmp).sProjCode), CStr(ZR1DS405.Cells(lngRow, 2))) = 0 Then
                        colO14(strTmp).sProjCode = colO14(strTmp).sProjCode & Chr(10) & CStr(ZR1DS405.Cells(lngRow, 2))
                    End If
                End If
            Next lngRow
            'Заполнение инв. проектов у заявки/позиции END
            
        End If
            

        If colO14.Count <> 0 Then
            Dim ReportRow As Integer: ReportRow = REP_STA_ROW
            'Dim SingleRow as clGKPZ
            'Set SingleRow = New clGKPZ
            Dim SingleRow As Variant
            For Each SingleRow In colO14
                wsReport.Cells(ReportRow, 21) = CStr(SingleRow.sPosition)
                wsReport.Cells(ReportRow, 22) = Replace(CStr(SingleRow.sLotID), "Не присвоено", "")
                wsReport.Cells(ReportRow, 1) = SingleRow.sRowPos
                wsReport.Cells(ReportRow, 3) = "'" & Replace(CStr(SingleRow.sProjCode), "#", "")
                wsReport.Cells(ReportRow, 2) = SingleRow.sTitleName
                wsReport.Cells(ReportRow, 4) = "" 'Пусто
                wsReport.Cells(ReportRow, 5) = "" 'Пусто
                wsReport.Cells(ReportRow, 6) = "" 'Пусто
                wsReport.Cells(ReportRow, 7) = Replace(CStr(SingleRow.sNetPower), "#", "")
                wsReport.Cells(ReportRow, 8) = "" 'Пусто
                wsReport.Cells(ReportRow, 9) = Replace(CStr(SingleRow.sNetPower), "#", "")
                wsReport.Cells(ReportRow, 10) = "" 'Пусто
                wsReport.Cells(ReportRow, 11) = "" 'Пусто
                wsReport.Cells(ReportRow, 12) = SingleRow.sName
                If CStr(SingleRow.sOrgBuy) = "Не присвоено" Then
                    If CStr(SingleRow.sOrgBuy2) <> "Не присвоено" Then
                        wsReport.Cells(ReportRow, 13) = CStr(SingleRow.sOrgBuy2)
                    End If
                Else
                    wsReport.Cells(ReportRow, 13) = CStr(SingleRow.sOrgBuy)
                End If
                
                'dblTemp = wsO14.Cells(lngRow, 30)
                'dblTemp = dblTemp '/ 1.18 dblTemp - 0.18 * dblTemp
                'wsReport.Cells(ReportRow, 14).Value = Round(dblTemp, 2) 'Format$(dblTemp, "###0.00")  '
                wsReport.Cells(ReportRow, 14).NumberFormat = "#,##0.00"
                'wsReport.Cells(ReportRow, 14) = VBA.Format$(SingleRow.sPlanPositionSum, "#,##0.00")
                wsReport.Cells(ReportRow, 14) = SingleRow.sPlanPositionSum
                wsReport.Cells(ReportRow, 14).HorizontalAlignment = xlLeft
                'wsReport.Cells(ReportRow, 14).Style = "Comma"
                wsReport.Cells(ReportRow, 15) = Replace(CStr(SingleRow.sAttachType), "Не присвоено", "")
                wsReport.Cells(ReportRow, 16) = CStr(SingleRow.sMethodBuy) ' Format$(dblTemp, "#,##0.0") '
                
                wsReport.Cells(ReportRow, 17) = Replace(CStr(SingleRow.sDateStartProcess), "#", "")
                wsReport.Cells(ReportRow, 18) = Replace(CStr(SingleRow.sDateCloseADeal), "#", "")
                wsReport.Cells(ReportRow, 19) = Replace(CStr(SingleRow.sPlanDate), "#", "")
                ReportRow = ReportRow + 1

            Next SingleRow

            RR = REP_STA_ROW + colO14.Count - 1
            'Очищаем область ниже полученных данных(область предыдущего построения)
            Set rngTemp = objRprtWS.Range(objRprtWS.Cells(REP_STA_ROW, REP_STA_COL), _
                                            objRprtWS.Cells(RR, REP_END_COL))
            Range_Format rngTemp
            Set rngTemp = Nothing

        End If

        Call Completion(objRprtWS)
        Application.StatusBar = "Формировании отчета завершено"
    End If
GoTo ends:
stop_macros:
    Application.StatusBar = "При формировании отчета возникла ошибка!"

ends:
    Set objDP = Nothing
    Set rngReport = Nothing
    Set objRprtWS = Nothing
    Set objBExWS = Nothing
End Sub
Private Function CaptionCon(ByRef vws As Worksheet, ByVal vRow As Long, ByVal vStart As Integer, ByVal vCount As Integer) As String
    Dim i As Integer
    Dim strTemp As String
    
    For i = 0 To vCount
        If vws.Cells(vRow, vStart + i) <> "#" Then
            strTemp = strTemp & vws.Cells(vRow, vStart + i)
        Else
            Exit For
        End If
    Next i
    CaptionCon = strTemp
End Function

'Пополнение коллекции датапровайдеров
Private Sub FillgcolDP(ParamArray varname())
    Dim dataRange        As Range
    Dim sideHeaderRange  As Range
    
    Dim objDPs As Object
    Set objDPs = BEx.DataProviders
    With gcolDP.Item(varname(0))
        'Получаем техническое имя запроса из свойств датапровайдера хранящихся в объекте BExApplication
        .Query = objDPs(varname(0)).Query

        'Определение начальной строки и столбца ячейки с данными относительно начала области вывода датапровайдера.
        .dataOffsetX = objDPs(varname(0)).Result.Grid.firstdatacell.x
        .dataOffsetY = objDPs(varname(0)).Result.Grid.firstdatacell.Y

        'Определение начальной строки и столбца области вывода датапровайдера относительно начала координат листа
        .DPOffsetX = varname(1).Column
        .DPOffsetY = varname(1).Row
        .isEmpty = False

        'Если данные в датапровайдере есть, данные - определение количества строк и столбцов с данными в таблице с результатом. Если значения свойств dataOffsetY и dataOffsetX >0, то либо датапровайдер не содержит данных, либо запрос не подержит показателей, например, построен на признаке и выводит список основных данных. Последняя ситуация будет обработана ниже:
        If .dataOffsetY > 0 And varname(1).Cells(1, 1) <> MSG_NO_VAR And varname(1).Cells(1, 1) <> MSG_NO_DATA Then
            .dataColumnsCount = varname(1).Columns.Count - .dataOffsetX
            .dataRowsCount = varname(1).Rows.Count - .dataOffsetY
            .isEmptyData = False
        Else
        'Отдельная обработка для датапровайдеров не содержащих показатели:
            .dataColumnsCount = varname(1).Columns.Count
            .dataRowsCount = varname(1).Rows.Count - 1
            .isEmptyData = True
            On Error Resume Next

            Set dataRange = varname(1).Offset(1, 0).Resize(varname(1).Rows.Count - 1)
            .dataAddress = dataRange.Address

            If Err.Number <> 0 Or varname(1).Cells(1, 1) = MSG_NO_VAR Or varname(1).Cells(1, 1) = MSG_NO_DATA Then
                .isEmpty = True
                On Error GoTo 0
            End If
        End If

        'Определение имени листа с данными:
        .dataSheetName = varname(1).Worksheet.Name

        If Not .isEmptyData Then
            'Определение области с данными:
            Set dataRange = Range(Cells(.DPOffsetY + .dataOffsetY, .DPOffsetX + .dataOffsetX), _
                                            Cells(.DPOffsetY + .dataOffsetY + .dataRowsCount - 1, .DPOffsetX + .dataOffsetX + .dataColumnsCount - 1))

            'Определение области с "боковиком":
            Set sideHeaderRange = Range(Cells(.DPOffsetY + .dataOffsetY, .DPOffsetX), _
                                            Cells(.DPOffsetY + .dataOffsetY + .dataRowsCount - 1, .DPOffsetX + .dataOffsetX - 1))
        End If
    End With
    
End Sub

'Создаем статическую функцию для подсчета количества Grid-ов. Для инициализации необходимо передать опциональные параметры toZero = True и myBEx = Bex (типа BExApplication).
Static Function GridCount(Optional toZero As Boolean, Optional ByRef myBEx As Variant) As Integer
    Dim locGridCount As Integer
    Dim myBExItem As Object
    
    If toZero Then
        locGridCount = 0
        For Each myBExItem In myBEx.Items
            If myBExItem.ToString Like "*BExItemGrid*" Then
                locGridCount = locGridCount + 1
            End If
        Next
    End If
    GridCount = locGridCount
    Set myBExItem = Nothing
End Function

Sub Range_Format(ByRef vRange As Range)
    vRange.VerticalAlignment = xlTop
    vRange.WrapText = True
'    With vRange
'        .HorizontalAlignment = xlGeneral
'        .VerticalAlignment = xlTop
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    With vRange
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlTop
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With

    With vRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With vRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With vRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With vRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With vRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With vRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub




