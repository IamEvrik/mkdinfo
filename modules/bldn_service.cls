VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ����� ������ �� ����
' ----------------------------------------------------------------------------
Private classBldnId As Long             ' ��� ����
Private classModeId As Long             ' ��� ������
Private classServiceId As Long          ' ��� ������
Private classInputsCount As Long        ' ���������� ������
Private classMode As service_mode       ' �����
Private classCanCounter As Boolean      ' ����������� ��������� ����
Private classNote As String             ' ����������


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ������������� - ����� ��������
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(BldnId As Long, serviceId As Long)
' ----------------------------------------------------------------------------
' ������������� ������ � ����
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_service_in_bldn"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnId").Value = BldnId
    cmd.Parameters("serviceId").Value = serviceId
    
    Set rst = cmd.Execute
    If rst.EOF Or rst.BOF Then GoTo errHandler
    
    classBldnId = BldnId
    classServiceId = serviceId
    classModeId = longValue(rst!mode_id)
    classInputsCount = longValue(rst!inputs_count)
    classCanCounter = rst!possible_counter
    classNote = DBgetString(rst!Note)
    
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


Property Get Service() As Service
' ----------------------------------------------------------------------------
' ������
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    If classBldnId <> NOTVALUE Then
        Set Service = services(CStr(classServiceId))
    End If
End Property


Property Get Mode() As service_mode
' ----------------------------------------------------------------------------
' �����
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    If classBldnId <> NOTVALUE And classModeId <> NOTVALUE Then
        If classMode Is Nothing Then
            Set classMode = New service_mode
            classMode.initial classModeId
        End If
        Set Mode = classMode
    End If
End Property


Property Get inputsCount() As Long
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    If classBldnId <> NOTVALUE Then
        inputsCount = classInputsCount
    End If
End Property


Property Get canCounter() As Boolean
' ----------------------------------------------------------------------------
' ����������� ��������� ����
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    If classBldnId <> NOTVALUE Then
        canCounter = classCanCounter
    End If
End Property


Property Get Note() As String
' ----------------------------------------------------------------------------
' ����������
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    If classBldnId <> NOTVALUE Then
        Note = classNote
    End If
End Property


Public Sub add(BldnId As Long, serviceId As Long, modeId As Long, _
                                        inputCounts As Long, _
                                        canCounter As Boolean, _
                                        Note As String)
' ----------------------------------------------------------------------------
' ���������� ������ ��� �������� � ����
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    classBldnId = BldnId
    classServiceId = serviceId
    classModeId = modeId
    classInputsCount = inputCounts
    classCanCounter = canCounter
    classNote = Note
End Sub


Public Sub create(BldnId As Long, serviceId As Long, modeId As Long, _
                                                    inputCounts As Long, _
                                                    canCounter As Boolean, _
                                                    Note As String)
' ----------------------------------------------------------------------------
' ���������� ������ � ����
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    Call update(BldnId:=BldnId, _
                serviceId:=serviceId, _
                modeId:=modeId, _
                inputCounts:=inputCounts, _
                canCounter:=canCounter, _
                Note:=Note, _
                addNew:=True)
End Sub


Public Sub update(BldnId As Long, serviceId As Long, modeId As Long, _
                                inputCounts As Long, canCounter As Boolean, _
                                Note As String, Optional addNew = False)
' ----------------------------------------------------------------------------
' ���������/���������� �������� � ����
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "add_bldn_service"
    Else
        cmd.CommandText = "change_bldn_service"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnid").Value = BldnId
    cmd.Parameters("serviceid").Value = serviceId
    cmd.Parameters("modeid").Value = modeId
    cmd.Parameters("inputs").Value = inputCounts
    cmd.Parameters("possCounter").Value = canCounter
    cmd.Parameters("newNote").Value = Note
    Set rst = cmd.Execute
    
    classBldnId = BldnId
    classServiceId = serviceId
    classModeId = modeId
    classInputsCount = inputCounts
    classCanCounter = canCounter
    classNote = Note
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_service.update", Err.Description
    End If
End Sub


Public Sub delete()
'-----------------------------------------------------------------------------
' �������� ������ �� ����
' Last update: 20.08.2018
'-----------------------------------------------------------------------------
    If classBldnId = NOTVALUE Or classServiceId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "bldn_service.delete", _
                                                            "������ �� �����"
    End If
    
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "delete_bldn_service"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnId").Value = classBldnId
    cmd.Parameters("serviceId").Value = classServiceId
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_service.delete", Err.Description
    End If
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' ����� ��������
' Last update: 20.08.2018
' ----------------------------------------------------------------------------
    classBldnId = NOTVALUE
    classServiceId = NOTVALUE
    classModeId = NOTVALUE
End Sub
