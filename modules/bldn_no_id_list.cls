VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_no_id_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список домов с минимальной информацией
' ----------------------------------------------------------------------------
Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set this = New Collection
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


Public Sub initial(Optional filterId As Long = ALLVALUES, _
                    Optional villageId As Long = ALLVALUES, _
                    Optional mdId As Long = ALLVALUES, _
                    Optional OnlyManaged As Boolean = False, _
                    Optional dogovorType As Long = ALLVALUES)
' ----------------------------------------------------------------------------
' заполнение коллекции начальными данными
' 27.09.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset, sqlParams As New Dictionary, sqlString As String
    Dim tempItem As bldn_no_id_class
    
    Set this = New Collection
    
    If OnlyManaged Then
        sqlString = "get_managed_bldn_id_no_list"
    Else
        sqlString = "get_bldn_id_no_list"
    End If
    
    sqlParams.add "InStreetId", filterId
    sqlParams.add "InVillageId", villageId
    sqlParams.add "InMdId", mdId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        If dogovorType = ALLVALUES Or dogovorType = rst!dogovor_type Then
            Set tempItem = New bldn_no_id_class
            tempItem.Address = rst!Address
            tempItem.Id = rst!bId
            tempItem.mcId = rst!mc_id
            tempItem.Name = rst!bldn_no
            tempItem.reportOut = rst!Out_Report
            tempItem.siteName = DBgetString(rst!site_name)
            tempItem.dogovorType = rst!dogovor_type
            this.add tempItem, CStr(tempItem.Id)
            Set tempItem = Nothing
        End If
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
End Sub


Public Sub initialWithAddress(Optional filterId As Long = ALLVALUES, _
                                Optional villageId As Long = ALLVALUES, _
                                Optional mdId As Long = ALLVALUES, _
                                Optional OnlyManaged As Boolean = False, _
                                Optional dogovorType As Long = ALLVALUES)
' ----------------------------------------------------------------------------
' заполнение коллекции начальными данными с указанием полного адреса
'        вместо номера дома
' 27.09.2022
' ----------------------------------------------------------------------------
    Dim i As Long
    
    Me.initial filterId, villageId, mdId, OnlyManaged, dogovorType
    For i = 1 To this.count
        this(i).Name = this(i).Address
    Next i
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As bldn_no_id_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции по индексу
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


