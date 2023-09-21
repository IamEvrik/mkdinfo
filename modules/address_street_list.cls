VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_street_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' About: список улиц
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' заполнение коллекции данными
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim cn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim tempItem As address_street_class
    
    Set this = New Collection
    
    Set cn = DBConnection.Connection
    Set rst = New ADODB.Recordset
    rst.Open "streetlist", cn, , adLockReadOnly, adCmdTable
    
    If rst.BOF And rst.EOF Then Exit Sub
    rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New address_street_class
        tempItem.add _
            streetId:=rst!Id, _
            StreetName:=DBgetString(rst!Name), _
            villageId:=rst!village_id, _
            streetTypeId:=rst!street_type, _
            siteName:=DBgetString(rst!site_name), _
            useFullName:=True
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set cn = Nothing
    Set rst = Nothing
End Sub


Public Sub initial(villageId As Long)
' ----------------------------------------------------------------------------
' заполнение коллекции данными по коду села
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Dim i As Long
    
    For i = this.count To 1 Step -1
        If this(i).Village.Id <> villageId Then this.remove i
    Next i
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 30.03.2016
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Long) As address_street_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 17.03.2016
' ----------------------------------------------------------------------------
    If itemidx > 0 And itemidx <= this.count Then
        Set Item = this.Item(itemidx)
    Else
        Err.Raise ERROR_INDEX_OUT_OF_DICT, "street_list.Item", _
                                            "Индекс за пределами справочника"
    End If
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 17.03.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Public Sub remove(itemidx As Long)
' ----------------------------------------------------------------------------
' удаление элемента из коллекции
' Last update: 17.03.2016
' ----------------------------------------------------------------------------
    this.remove itemidx
End Sub
