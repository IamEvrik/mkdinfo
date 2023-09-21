VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plan_statuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список статусов планируемых работ (статический)
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 12.02.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As plan_status
    Dim sqlParams As New Dictionary, sqlString As String
    
    Set this = New Collection
    
    sqlString = "get_plan_statuses"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New plan_status
        tempItem.Id = rst!Id
        tempItem.Name = rst!Name
        tempItem.inPlan = rst!in_plan
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set sqlParams = Nothing
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 18.07.2017
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
' Last update: 15.03.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As plan_status
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get PWPlan() As Long
' ----------------------------------------------------------------------------
' код статуса "В плане"
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    PWPlan = 1
End Property


Property Get PWDone() As Long
' ----------------------------------------------------------------------------
' код статуса "Выполнен"
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    PWDone = 3
End Property


Property Get PWInWork() As Long
' ----------------------------------------------------------------------------
' код статуса "В работе"
' Last update: 09.08.2018
' ----------------------------------------------------------------------------
    PWInWork = 2
End Property

