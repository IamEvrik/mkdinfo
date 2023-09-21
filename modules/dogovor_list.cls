VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dogovor_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список видов договоров (статический)
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
    Dim tempItem As dogovor_class
    
    Set this = New Collection
    
    Set cn = DBConnection.Connection
    Set rst = New ADODB.Recordset
    rst.Open DogovorTableName, cn, , adLockReadOnly, adCmdTable
    
    If rst.BOF And rst.EOF Then Exit Sub
    rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New dogovor_class
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

Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As dogovor_class
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
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    count = this.count
End Property


Public Function remove(itemidx As Long) As Long
' ----------------------------------------------------------------------------
' удаление элемента из коллекции
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    If itemidx > 0 And itemidx < this.count Then
        this.remove itemidx
    End If
End Function
