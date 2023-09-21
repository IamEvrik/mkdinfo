VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_village_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' About: список населённых пунктов (статический)
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Call reload
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' получение всех населённых пунктов
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As address_village_class
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_villages"
    cmd.CommandType = adCmdStoredProc
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New address_village_class
        tempItem.add itemId:=rst!Id, _
                    municipalDistrictId:=rst!md_id, _
                    villageName:=rst!Name, _
                    siteName:=DBgetString(rst!site_name)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "villages.initial", Err.Description
    End If
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As address_village_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 12.02.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property

