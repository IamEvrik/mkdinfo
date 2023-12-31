VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "id_name_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ������ �� ������� ��� � �������� (�� �� ����)
' ----------------------------------------------------------------------------

Private this As Collection


Public Sub initGas()
' ----------------------------------------------------------------------------
' ��� �������������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmp As id_name_class
    Dim i As Long
    
    Set this = New Collection
    
    Set tmp = New id_name_class
    tmp.Id = 0
    tmp.Name = "�����������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 1
    tmp.Name = "�������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 2
    tmp.Name = "���������"
    this.add tmp, CStr(tmp.Id)
    
    Set tmp = Nothing
End Sub


Public Sub initHeating()
' ----------------------------------------------------------------------------
' ��� ���������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmp As id_name_class
    Dim i As Long
    
    Set this = New Collection
    
    Set tmp = New id_name_class
    tmp.Id = 0
    tmp.Name = "�����������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 1
    tmp.Name = "�����������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 2
    tmp.Name = "��������������"
    this.add tmp, CStr(tmp.Id)
    
    Set tmp = Nothing
End Sub


Public Sub initHotWater()
' ----------------------------------------------------------------------------
' ��� �������� �������������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmp As id_name_class
    Dim i As Long
    
    Set this = New Collection
    
    Set tmp = New id_name_class
    tmp.Id = 0
    tmp.Name = "�����������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 1
    tmp.Name = "�����������"
'    tmp.Name = "��������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 2
    tmp.Name = "�����������"
'    tmp.Name = "��������"
    this.add tmp, CStr(tmp.Id)
    
    Set tmp = Nothing
End Sub


Public Sub initColdWater()
' ----------------------------------------------------------------------------
' ��� ��������� �������������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmp As id_name_class
    Dim i As Long
    
    Set this = New Collection
    
    Set tmp = New id_name_class
    tmp.Id = 0
    tmp.Name = "�����������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = 1
    tmp.Name = "�����������"
    this.add tmp, CStr(tmp.Id)

    Set tmp = Nothing
End Sub


Public Sub initYesNo()
' ----------------------------------------------------------------------------
' ��/���
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    Dim tmp As id_name_class
    Dim i As Long
    
    Set this = New Collection
    
    Set tmp = New id_name_class
    tmp.Id = CLng(False)
    tmp.Name = "���"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = CLng(True)
    tmp.Name = "��"
    this.add tmp, CStr(tmp.Id)

    Set tmp = Nothing
End Sub


Public Sub initYesNoOther(otherString As String)
' ----------------------------------------------------------------------------
' ��/���/������
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    Dim tmp As id_name_class
    Dim i As Long
    
    Set this = New Collection
    
    Set tmp = New id_name_class
    tmp.Id = CLng(False)
    tmp.Name = "���"
    this.add tmp, CStr(tmp.Id)

    Set tmp = New id_name_class
    tmp.Id = CLng(True)
    tmp.Name = "��"
    this.add tmp, CStr(tmp.Id)
    
    Set tmp = New id_name_class
    tmp.Id = OTHERVALUE
    tmp.Name = otherString
    this.add tmp, CStr(tmp.Id)

    Set tmp = Nothing
End Sub


Public Sub initBldnExpenseName(expenseId As Long, nameId As Long)
' ----------------------------------------------------------------------------
' ���� � ��������� ������ ��������
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "bldn_list_use_expense_name"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("expId").Value = expenseId
    cmd.Parameters("expname").Value = nameId
    Set rst = cmd.Execute
    
    If Not (rst.BOF Or rst.EOF) Then
    
        Do While Not rst.EOF
            Set tempItem = New id_name_class
            tempItem.Id = rst!BldnId
            tempItem.Name = DBgetString(rst!address)
            
            this.add tempItem, CStr(tempItem.Id)
            Set tempItem = Nothing
            rst.MoveNext
        Loop
    End If
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set cmd = Nothing
    Set rst = Nothing
End Sub


Public Sub initBldnTypes()
' ----------------------------------------------------------------------------
' ���� ����
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_bldn_types"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("userId").Value = CurrentUser.userId
    Set rst = cmd.Execute
    
    If Not (rst.BOF Or rst.EOF) Then
    
        Do While Not rst.EOF
            Set tempItem = New id_name_class
            tempItem.Id = rst!Id
            tempItem.Name = DBgetString(rst!Name)
            
            this.add tempItem, CStr(tempItem.Id)
            Set tempItem = Nothing
            rst.MoveNext
        Loop
    End If
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set cmd = Nothing
    Set rst = Nothing
End Sub


Public Sub initAccessTypes()
' ----------------------------------------------------------------------------
' ���� �������
' Last update: 20.09.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "adm_get_access_types"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("userId").Value = CurrentUser.userId
    Set rst = cmd.Execute
    
    If Not (rst.BOF Or rst.EOF) Then
    
        Do While Not rst.EOF
            Set tempItem = New id_name_class
            tempItem.Id = rst!Id
            tempItem.Name = DBgetString(rst!Name)
            
            this.add tempItem, CStr(tempItem.Id)
            Set tempItem = Nothing
            rst.MoveNext
        Loop
    End If
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set cmd = Nothing
    Set rst = Nothing
End Sub


Public Sub initRoleAccess(roleId As Long, accessType As Long, _
                                                            isHas As Boolean)
' ----------------------------------------------------------------------------
' ������ � ����
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If isHas Then
        cmd.CommandText = "adm_role_has_access"
    Else
        cmd.CommandText = "adm_role_has_no_access"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("userId").Value = CurrentUser.userId
    cmd.Parameters("acsType").Value = accessType
    cmd.Parameters("roleId").Value = roleId
    Set rst = cmd.Execute
    
    If Not (rst.BOF Or rst.EOF) Then
    
        Do While Not rst.EOF
            Set tempItem = New id_name_class
            tempItem.Id = rst!Id
            tempItem.Name = DBgetString(rst!Name)
            
            this.add tempItem, CStr(tempItem.Id)
            Set tempItem = Nothing
            rst.MoveNext
        Loop
    End If
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set cmd = Nothing
    Set rst = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' ���������� ��������� � ���������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As id_name_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ��������� �� �������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' ----------------------------------------------------------------------------
' ��� ������ ��������� � ������� For each ...
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Set NewEnum = this.[_NewEnum]
End Function
