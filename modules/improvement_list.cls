VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "improvement_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список классов режимов потребления
' ----------------------------------------------------------------------------
Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' заполнение списка
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim tempItem As improvement_class
    
    Set this = New Collection
    
    Set cn = DBConnection.Connection
    Set rst = New ADODB.Recordset
    rst.Open ImprovementTableName, cn, , adLockReadOnly, adCmdTable
    
    If rst.BOF And rst.EOF Then Exit Sub
    rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New improvement_class
        tempItem.add _
            Id:=rst!Id, _
            FullName:=rst!Name, _
            ShortName:=DBgetString(rst!short_name)
        
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set cn = Nothing
    Set rst = Nothing
End Sub


Property Get Item(itemidx As Variant) As improvement_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 28.03.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property
