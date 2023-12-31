VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "globalWorkType_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ����� ���� ������� (�������, �����������, ����������)
' ----------------------------------------------------------------------------

Private classId As Long             ' ���
Private className As String         ' ��������
Private classNote As String         ' ����������


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ����������� ����� �������������
' Last update: 16.03.2016
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
    cmd.CommandText = "getGWT"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = DBgetString(rst!Name)
    classNote = DBgetString(rst!Description)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "gwt_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Integer
' ----------------------------------------------------------------------------
' ��� ������ ������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' �������� ������ ������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get Note() As String
' ----------------------------------------------------------------------------
' ���������� � ������ ������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Note = classNote
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' ��������� �������� ������ ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=newValue)
End Property


Property Let Note(newValue As String)
' ----------------------------------------------------------------------------
' ��������� ���������� � ������ ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newNote:=newValue)
End Property


Public Sub add(Id As Long, Name As String, Note As String)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ����������
' Last update: 12.04.2016
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
    classNote = Note
End Sub


Public Sub create(Name As String, Note As String)
' ----------------------------------------------------------------------------
' ����������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newNote:=Note, addNew:=True)
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
    cmd.CommandText = "deleteGWT"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "gwt_class.delete", _
                                            "���������� ������� ��� �������"
        Else
            Err.Raise Err.Number, "gwt_class.delete", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional newNote As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' ��������� (����������) �������� � ����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "gwt_class.update", "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    ' ���� ��������� �� ������, �� �� ������������� �������� ������
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    If StrComp(newNote, NOTSTRING) = 0 Then newNote = classNote
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createGWT"
    Else
        cmd.CommandText = "changeGWT"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 50, newName)
    cmd.Parameters.Append cmd.CreateParameter("newdesc", adVarChar, _
                                adParamInput, Len(newNote) + 1, newNote)
    
    Set rst = cmd.Execute
    
    className = newName
    classNote = newNote
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
    
    If Err.Number <> 0 Then
        If errorNotUnique(Err.Description) Then
            Err.Raise ERROR_NOT_UNIQUE, "gwt_class.update", _
                                    "����� ��� ������� ��� ���� � �����������"
        Else
            Err.Raise Err.Number, "gwt_class.update", Err.Description
        End If
    End If
        
    Call reloadList
End Sub

Private Sub reloadList()
' ----------------------------------------------------------------------------
' ���������� ������ ����� ��������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    globalWorkType_list.reload
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' ��������� ��������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
