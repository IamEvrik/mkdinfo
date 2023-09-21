VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' услуги в доме
' ----------------------------------------------------------------------------

Private this As Collection


Public Sub initial(BldnId As Long)
' ----------------------------------------------------------------------------
' инициализация
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As bldn_service
    
    On Error GoTo errHandler
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_bldn_services"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnId").Value = BldnId
    
    Set rst = cmd.Execute
    If rst.EOF Or rst.BOF Then GoTo errHandler
    
    Do While Not rst.EOF
        Set tempItem = New bldn_service
        tempItem.add _
                    BldnId:=rst!bldn_id, _
                    serviceId:=rst!service_id, _
                    modeId:=longValue(rst!mode_id), _
                    inputCounts:=longValue(rst!inputs_count), _
                    canCounter:=rst!possible_counter, _
                    Note:=DBgetString(rst!Note)
                    
        this.add tempItem
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_services.initial", Err.Description
    End If
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 02.07.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 02.07.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As bldn_service
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 02.07.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


