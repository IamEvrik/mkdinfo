VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "employee_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------
' ��������
'-----------------------------------------------------------------------------
Private classId As Long                             ' ���
Private classUK As Long                             ' ����������� ��������
Private classLastName As String                     ' �������
Private classFirstName As String                    ' ���
Private classSecondName As String                   ' ��������
Private classPosition As String                     ' ���������
Private classIsReport As Boolean                    ' ������������� � ��������
Private classPositionStatus As Long                 ' ���������
    
    
Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ������������� ������� ������
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Call flushValues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' ������������� ������ �� ����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If ItemId = NOTVALUE Then Exit Sub

    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getEmployee"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    ItemId)
    
    Set rst = cmd.Execute
    
    classId = ItemId
    classFirstName = DBgetString(rst!first_name)
    classSecondName = DBgetString(rst!second_name)
    classLastName = DBgetString(rst!last_name)
    classUK = rst!organization_id
    classPositionStatus = rst!position_status
    classPosition = DBgetString(rst!position_name)
    classIsReport = rst!sign_report
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "employee_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' ���
' Last update: 14.03.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get uk() As uk_class
' ----------------------------------------------------------------------------
' ����������� ��������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set uk = uk_class(CStr(classUK))
End Property


Property Get LastName() As String
' ----------------------------------------------------------------------------
' �������
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    LastName = classLastName
End Property


Property Get FirstName() As String
' ----------------------------------------------------------------------------
' ���
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    FirstName = classFirstName
End Property


Property Get SecondName() As String
' ----------------------------------------------------------------------------
' ��������
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    SecondName = classSecondName
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' ��� (��� ������ � ������)
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Name = FIO
End Property


Property Get FullName() As String
' ----------------------------------------------------------------------------
' ������ ���
' 02.10.2022
' ----------------------------------------------------------------------------
    FullName = classLastName & " " & classFirstName & " " & classSecondName
End Property


Property Get FIO() As String
' ----------------------------------------------------------------------------
' ���
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        FIO = NOTSTRING
    Else
        FIO = classLastName + " " + Left(classFirstName, 1) + ". " + _
                                                Left(classSecondName, 1) + "."
    End If
End Property


Property Get Signature() As Byte()
' ----------------------------------------------------------------------------
' ������� ��������� � ���� ������� ����
' 04.10.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    If classId = NOTVALUE Then Exit Property
    
    sqlString = "get_employee_signature"
    sqlParams.add "InEmployeeId", classId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF Or rst.BOF Then GoTo errHandler
    
    If Not IsNull(rst!Signature) Then Signature = rst!Signature
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "employee_class.signature", Err.Description
    End If
End Property


Property Get Position() As String
' ----------------------------------------------------------------------------
' ���������
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    Position = classPosition
End Property


Property Get PositionStatus() As Long
' ----------------------------------------------------------------------------
' ���� ���������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    PositionStatus = classPositionStatus
End Property


Property Get ReportSign() As Boolean
' ----------------------------------------------------------------------------
' ����������� �� ������� ����������
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    ReportSign = classIsReport
End Property


Property Let setMcId(newValue As Long)
' ----------------------------------------------------------------------------
' ��������� �� �� ����
' Last update: 27.02.2018
' ----------------------------------------------------------------------------
    Call update(newMC:=newValue)
End Property


Property Let LastName(newValue As String)
' ----------------------------------------------------------------------------
' ��������� �������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newLastName:=newValue)
End Property


Property Let FirstName(newValue As String)
' ----------------------------------------------------------------------------
' ��������� �����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newFirstName:=newValue)
End Property


Property Let SecondName(newValue As String)
' ----------------------------------------------------------------------------
' ��������� ��������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newSecondName:=newValue)
End Property


Property Let Position(newValue As String)
' ----------------------------------------------------------------------------
' ��������� ���������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newPosition:=newValue)
End Property


Property Let ReportSign(newValue As Boolean)
' ----------------------------------------------------------------------------
' ��������� ����� ������� ��������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newSign:=newValue)
End Property


Property Let PositionStatus(newValue As Long)
' ----------------------------------------------------------------------------
' ��������� ���� ���������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newPositionStatus:=newValue)
End Property


Public Sub add(Id As Long, orgId As Long, LastName As String, _
                FirstName As String, SecondName As String, _
                Position As String, isSign As Boolean, _
                PositionStatus As Long)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ����������
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    classId = Id
    classUK = orgId
    classLastName = LastName
    classFirstName = FirstName
    classSecondName = SecondName
    classPosition = Position
    classIsReport = isSign
    classPositionStatus = PositionStatus
End Sub


Public Sub create(FirstName As String, orgId As Long, SecondName As String, _
                LastName As String, Position As String, isSign As Boolean, _
                PositionStatus As Long)
' ----------------------------------------------------------------------------
' ���������� ������ ���������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newFirstName:=FirstName, newSecondName:=SecondName, _
                newLastName:=LastName, newMC:=orgId, newPosition:=Position, _
                newSign:=isSign, newPositionStatus:=PositionStatus, _
                addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' �������� ���������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "employee_class.delete", _
                                                            "������ �� �����"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteEmployee"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "employee_class.delete", _
                                            "���������� ������� ����������"
        Else
            Err.Raise Err.Number, "employee_class.delete", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Public Sub update(Optional newFirstName As String = NOTSTRING, _
                    Optional newSecondName As String = NOTSTRING, _
                    Optional newLastName As String = NOTSTRING, _
                    Optional newMcId As Long = NOTVALUE, _
                    Optional newPosition As String = NOTSTRING, _
                    Optional newSign As Variant = Nothing, _
                    Optional newPositionStatus As Long = NOTVALUE, _
                    Optional addNew As Boolean = False)
'-----------------------------------------------------------------------------
' ���������� ���������� � ����
' Last update: 28.03.2018
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' ���� ��������� �� ������, �� �� ������������� �������� ������
    If StrComp(newFirstName, NOTSTRING) = 0 Then newFirstName = classFirstName
    If StrComp(newSecondName, NOTSTRING) = 0 Then newSecondName = classSecondName
    If StrComp(newLastName, NOTSTRING) = 0 Then newLastName = classLastName
    If StrComp(newPosition, NOTSTRING) = 0 Then newPosition = classPosition
    If newMcId = NOTVALUE Then newMcId = classUK
    If newPositionStatus = NOTVALUE Then newPositionStatus = classPositionStatus
    If IsMissing(newSign) Then newSign = classIsReport
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createEmployee"
    Else
        cmd.CommandText = "changeEmployee"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newfname", adVarChar, _
                                            adParamInput, 60, newFirstName)
    cmd.Parameters.Append cmd.CreateParameter("newsname", adVarChar, _
                                            adParamInput, 60, newSecondName)
    cmd.Parameters.Append cmd.CreateParameter("newlname", adVarChar, _
                                            adParamInput, 60, newLastName)
    cmd.Parameters.Append cmd.CreateParameter("neworgid", adUnsignedInt, _
                                            adParamInput, , newMcId)
    cmd.Parameters.Append cmd.CreateParameter("newpos", adUnsignedInt, _
                                            adParamInput, , newPositionStatus)
    cmd.Parameters.Append cmd.CreateParameter("newposname", adVarChar, _
                                            adParamInput, 100, newPosition)
    cmd.Parameters.Append cmd.CreateParameter("newsignreport", adBoolean, _
                                            adParamInput, , newSign)
    Set rst = cmd.Execute
    
    classFirstName = newFirstName
    classSecondName = newSecondName
    classLastName = newLastName
    classUK = newMcId
    classPosition = newPosition
    classIsReport = newSign
    classPositionStatus = newPositionStatus
    If addNew Then classId = rst!newId
        
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
            Err.Raise ERROR_NOT_UNIQUE, "employee_class.update", _
                                "� ����������� ����� ���� ������ ���� " & _
                                "�������� ��� ������� �������"
        Else
            Err.Raise Err.Number, "employee_class.update", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' ���������� ������ ����������� �����������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    uk_list(CStr(classUK)).initialEmployees
End Sub


Private Sub flushValues()
' ----------------------------------------------------------------------------
' ��������� ���� ��������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
