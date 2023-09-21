VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plan_terms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' даты для планов
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - ничего не делаем
' Last update: 17.05.2019
' ----------------------------------------------------------------------------

End Sub


Public Sub reload(beginDate As Date, yearsCount As Long)
' ----------------------------------------------------------------------------
' Заполнение списка с даты beginDate включительно на yearsCount годов
' Last update: 17.05.2019
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim tmpItem As id_name_class
    Dim tmpDate As Date
    
    Set this = New Collection
    
    For i = 1 To yearsCount * 12
        tmpDate = DateAdd("m", i - 1, beginDate)
        Set tmpItem = New id_name_class
        tmpItem.Id = i
        tmpItem.Name = MonthName(Month(tmpDate)) & " " & Year(tmpDate)
        this.add tmpItem, CStr(tmpItem.Id)
        Set tmpItem = Nothing
    Next i
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 17.05.2019
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 17.05.2019
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As id_name_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 17.05.2019
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get IdByDate(InDate As Date) As Long
' ----------------------------------------------------------------------------
' код элемента по дате
' Last update: 31.08.2020
' ----------------------------------------------------------------------------
    Dim curItem As id_name_class
    IdByDate = NOTVALUE
    For Each curItem In this
        If curItem.Name = MonthName(Month(InDate)) & " " & Year(InDate) Then
            IdByDate = curItem.Id
            Exit For
        End If
    Next
End Property
