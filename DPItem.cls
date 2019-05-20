VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DPItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DPName        As String '– Техническое имя датапровайдера
Public Query         As String '– Техническое имя запроса
Public dataSheetName As String '– Имя листа с данными
Public sideHeaderAddress  As String '– адрес области с "боковиком", "урезанный" до одного столбца
Public sideHeaderAddressFull   As String '– адрес области с "боковиком", "полный"
Public dataAddress   As String '– адрес области с данными
Public localHeaderAddress As String '– адрес заголовка столбцов. включая строку с единицами измерения
Public isEmpty       As Boolean '– Есть данные в провайдере данных
Public isEmptyData   As Boolean '– Есть данные на показателях в провайдере данных
Public isUpdated     As Boolean '– Метаданные провайдера обновлены
Public DPOffsetX     As Integer '– X - смещение первой ячейки датапровайдера относительно начала координат листа
Public DPOffsetY     As Integer '– Y - смещение первой ячейки датапровайдера относительно начала координат листа
Public dataOffsetX   As Integer '– Х - смещение первой ячейки с данными относительно начала датапровайдера
Public dataOffsetY   As Integer '– Y - смещение первой ячейки с данными относительно начала датапровайдера
Public dataColumnsCount   As Integer '– Количество столбцов с данными в датапровайдере
Public dataRowsCount As Long '– Количество строк с данными в датапровайдере

Public Property Get FirstRow() As Integer
   FirstRow = DPOffsetY + IIf(dataOffsetY < 0, 0, dataOffsetY - 1) + 1
End Property

Public Property Get LastRow() As Integer
    LastRow = dataRowsCount + DPOffsetY + IIf(dataOffsetY < 0, 0, dataOffsetY - 1)
End Property

Public Property Get FirstCol() As Integer
   FirstCol = 1 + 1
End Property

Public Property Get FirstDataCol() As Integer
   FirstDataCol = DPOffsetX + dataOffsetX
End Property

Public Property Get LastDataCol() As Integer
    LastDataCol = DPOffsetX + dataOffsetX + dataColumnsCount - 1
End Property


