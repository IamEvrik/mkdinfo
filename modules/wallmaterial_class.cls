VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wallmaterial_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' �������� ����
' ----------------------------------------------------------------------------

Private classId As Long             ' ���
Private className As String         ' ��������


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ����������� ����� �������������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getWallMaterial"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = DBgetString(rst!Name)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "wallmaterial_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' ���
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' ��������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' ��������� ��������
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=newValue)
End Property


Public Sub add(Id As Long, Name As String)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ����������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
End Sub


Public Sub create(Name As String)
' ----------------------------------------------------------------------------
' ����������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call update(newName:=Name, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' ��������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "wallmaterial_class.delete", _
                                                            "������ �� �����"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteWallMaterial"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "wallmaterial_class.delete", _
                                            "���������� ������� �������� ����"
        Else
            Err.Raise Err.Number, "wallmaterial_class.delete", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' ��������� (����������) �������� � ����
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "wallmaterial_class.update", _
                                                            "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    ' ���� ��������� �� ������, �� �� ������������� �������� ������
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createWallMaterial"
    Else
        cmd.CommandText = "changeWallMaterial"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 100, newName)
    
    Set rst = cmd.Execute
    
    className = newName
    If addNew Then classId = rst!newid
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "wallmaterial_class.update", Err.Description
        
    Call reloadList
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' ��������� ��������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    wallmaterial_list.reload
End Sub


