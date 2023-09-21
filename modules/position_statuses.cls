VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "position_statuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список должностей (статический)
' ----------------------------------------------------------------------------

Private this As Collection

Private Enum statusesId
    directorId = 1
    chiefId = 2
    otherID = 3
End Enum


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim tempItem As plan_status
    
    Set this = New Collection
    
    Set cn = DBConnection.Connection
    Set rst = New ADODB.Recordset
    rst.Open PositionStatusTableName, cn, , adLockReadOnly, adCmdTable
    
    If rst.BOF And rst.EOF Then Exit Sub
    rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New plan_status
        tempItem.Id = rst!Id
        tempItem.Name = rst!Name
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set cn = Nothing
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 27.03.2018
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
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As position_status
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get Director() As Long
' ----------------------------------------------------------------------------
' код должности директора
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    Director = statusesId.directorId
End Property


Property Get ChiefIngeneer() As Long
' ----------------------------------------------------------------------------
' код должности главного инженера
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    ChiefIngeneer = statusesId.chief
End Property


Property Get Other() As Long
' ----------------------------------------------------------------------------
' код должности других должностей
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    Other = statusesId.otherID
End Property

