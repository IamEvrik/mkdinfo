VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "village_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список классов типов населённых пунктов (статический, пока не в базе)
' ----------------------------------------------------------------------------
Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 23.03.2018
' ----------------------------------------------------------------------------
    Dim cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim tempItem As village_type
    
    Set this = New Collection
    
    Set cn = DBConnection.Connection
    Set rst = New ADODB.Recordset
    rst.Open VillageTypeTableName, cn, , adLockReadOnly, adCmdTable
    
    If rst.BOF And rst.EOF Then Exit Sub
    rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New village_type
        tempItem.Id = rst!Id
        tempItem.Name = rst!Name
        tempItem.ShortName = rst!short_name
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
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As village_type
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property
