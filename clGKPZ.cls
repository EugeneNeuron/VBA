VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clGKPZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public sInvProjName As String 'Наименование инв.проекта (стлб. 2)
'Public sInvProjID As String 'Идентификатор инв.проекта (стлб. 3)
'Public sPurchProd As String 'Наименование закупаемой продукции (стлб 12)
'Public sOrg As String 'Организатор закупки (стлб 13)
'Public sPlPrice As String 'Планируемая цена закупки (стлб 14)
'Public sPurchMethod As String 'Планируемый способ закупки (стлб 16)
'Public sNoticePlDate As String 'Планируемая дата объявления ЗП (стлб 17)
'Public sSignPlDate As String 'Планируема дата заключения договора (стлб 18)
'Public sStartPlDate As String 'Планируемая дата начала поставки/работ (стлб 19)
Public sPosition As String 'Заявка.Позиция (POS_ID = 4 в 21)
Public sLotID As String 'Лот ЗП (LOT_ID = 3 в 22)
Public sProjCode As String 'Код проекта (из 10 в 3)
Public sTitleName As String 'Наименование титула (из 20-27 в 2)
Public sNetPower As String 'ПроектМощПротяжСети1 (из 19 в 7,9)
Public sName As String 'Наименование (из 12 в 12)
Public sOrgBuy As String 'Организатор закупки (из 7 в 13)
Public sOrgBuy2 As String 'Организатор закупки (из 2 в 13)
Public sPlanPositionSum As String 'План сумма позиции (из 30 в 14)
Public sAttachType As String 'Тип прикрепленного документа (из 5 в 15)
Public sMethodBuy As String 'Способ закупки (из 6 в 16)
Public sDateStartProcess As String 'ДатаОбъявлОНачПроц (из 28 в 17)
Public sDateCloseADeal As String 'ПланДатаЗаключДог (из 11 в 18)
Public sPlanDate As String 'ПланДтНачалаПоставки (из 8 в 19)
Public sRowPos As String 'Номер отображенной строки (1)




