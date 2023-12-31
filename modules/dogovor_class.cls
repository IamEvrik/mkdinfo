VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dogovor_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ��������
' ----------------------------------------------------------------------------

Private classId As Long             ' ���
Private classFullName As String     ' ��������
Private classShortName As String    ' �������� ��������


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ����������� ����� �������������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getDogovor"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    classFullName = DBgetString(rst!Name)
    classShortName = DBgetString(rst!short_name)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "dogovor_class.initial", Err.Description
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
    Name = classShortName
End Property


Property Get FullName() As String
' ----------------------------------------------------------------------------
' ������ ��������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    FullName = classFullName
End Property


Property Get ShortName() As String
' ----------------------------------------------------------------------------
' ������� ��������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    ShortName = classShortName
End Property


Property Let FullName(newValue As String)
' ----------------------------------------------------------------------------
' About: ��������� ������� ��������
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "dogovor_class.fullname_let", _
                                                            "������ �� �����"
    End If

    Call update(newFullName:=newValue)
End Property


Property Let ShortName(newValue As String)
' ----------------------------------------------------------------------------
' ��������� �������� ��������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "dogovor_class.shortname_let", _
                                                            "������ �� �����"
    End If

    Call update(newShortName:=newValue)
End Property


Public Sub add(Id As Long, FullName As String, ShortName As String)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ����������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    classId = Id
    classFullName = FullName
    classShortName = ShortName
End Sub


Public Sub create(FullName As String, ShortName As String)
' ----------------------------------------------------------------------------
' ����������
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call update(newFullName:=FullName, newShortName:=ShortName, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' ��������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "dogovor_class.delete", _
                                                            "������ �� �����"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteDogovor"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "dogovor_class.delete", _
                        "���������� ������� �������, �� ������� ���� ������"
        Else
            Err.Raise Err.Number, "dogovor_class.delete", Err.Description
        End If
    End If
    
    Call reloadList

End Sub


Public Sub update(Optional newFullName As String = NOTSTRING, _
                        Optional newShortName As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' ��������� (����������) �������� � ����
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' ���� ��������� �� ������, �� �� ������������� �������� ������
    If StrComp(newFullName, NOTSTRING) = 0 Then newFullName = classFullName
    If StrComp(newShortName, NOTSTRING) = 0 Then newShortName = classShortName
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createDogovor"
    Else
        cmd.CommandText = "changeDogovor"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 100, newFullName)
    cmd.Parameters.Append cmd.CreateParameter("newshortname", adVarChar, _
                                            adParamInput, 30, newShortName)
    
    Set rst = cmd.Execute
    
    classFullName = newFullName
    classShortName = newShortName
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
        Err.Raise Err.Number, "dogovor_class.update", Err.Description
        
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
    dogovor_list.reload
End Sub
