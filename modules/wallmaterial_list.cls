VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wallmaterial_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список материалов стен
' ----------------------------------------------------------------------------
Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call reload
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' заполнение коллекции списком
' Last update: 04.04.2018
' ----------------------------------------------------------------------------
    Dim cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim tempItem As wallmaterial_class
    
    Set this = New Collection
    
    Set cn = DBConnection.Connection
    Set rst = New ADODB.Recordset
    rst.Open WallMaterialTableName, cn, , adLockReadOnly, adCmdTable
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New wallmaterial_class
        tempItem.add Id:=rst!Id, _
                        Name:=rst!Name
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
        
cleanHandler:
    rst.Close
    Set cn = Nothing
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


Property Get Item(itemidx As Variant) As wallmaterial_class
Attribute Item.VB_UserMemId = 0
' -----------------------------------------------------------------------------
' элемент коллекции
' Last update: 25.03.2018
' -----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
