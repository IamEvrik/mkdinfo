VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_village_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ��������� �����
' ----------------------------------------------------------------------------

Private classId As Long             ' ���
Private className As String         ' ��������
Private classMD As Long             ' ������������� �����������
Private classSite As String         ' ���� �� �����


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_village"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemid").Value = itemId
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = rst!Name
    classMD = rst!md_id
    classSite = rst!site_name
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "village_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' ���
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' ��������
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get Site() As String
' ----------------------------------------------------------------------------
' �������� ��� �����
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Site = classSite
End Property


Property Get Municipal_district() As address_md_class
' ----------------------------------------------------------------------------
' ������������� �����������
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Set Municipal_district = address_md_list(CStr(classMD))
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' ��������� �������� ���������� ������
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.name_let", _
                                                            "������ �� �����"
    End If
    
    Call update(newName:=newValue)
End Property


Property Let Site(newValue As String)
' ----------------------------------------------------------------------------
' ��������� �������� ��� �����
' Last update: 05.05.2016
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.site_let", _
                                                            "������ �� �����"
    End If
    
    Call update(newSite:=newValue)
End Property


Property Let setMd(newValue As Long)
' ----------------------------------------------------------------------------
' ��������� ���� �������������� �����������
' Last update: 05.05.2016
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.mdid_let", _
                                                            "������ �� �����"
    End If
    
    Call update(newMd:=newValue)
End Property


Public Sub add(itemId As Long, municipalDistrictId As Long, _
                                    villageName As String, siteName As String)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ���������� � ����
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    classId = itemId
    className = villageName
    classMD = municipalDistrictId
    classSite = siteName
End Sub


Public Sub create(villageName As String, municipalDistrictId As Long, _
                                                         siteName As String)
' ----------------------------------------------------------------------------
' ���������� ���������� ������
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=villageName, _
                newMd:=municipalDistrictId, _
                newSite:=siteName, _
                addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' �������� ���������� ������
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.delete", _
                                                            "������ �� �����"
    End If

    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "delete_village"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = classId
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
        
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "village_class.delete", _
                            "���������� ������� ����, � ������� ���� �����"
        Else
            Err.Raise Err.Number, "village_class.delete", Err.Description
        End If
    End If
    
    Call reloadVillageList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional newMd As Long = NOTVALUE, _
                        Optional newSite As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' ��������� (����������) �������� � ����
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' ���� ��������� �� ������, �� �� ������������� �������� ������
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    If newMd = NOTVALUE Then newMd = classMD
    If StrComp(newSite, NOTSTRING) = 0 Then newSite = classSite
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "create_village"
    Else
        cmd.CommandText = "change_village"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters("newname").Value = newName
    cmd.Parameters("newmd").Value = newMd
    cmd.Parameters("newsite").Value = newSite
    If Not addNew Then cmd.Parameters("itemId").Value = classId
    
    Set rst = cmd.Execute
    
    className = newName
    classMD = newMd
    classSite = newSite
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
            Err.Raise ERROR_NOT_UNIQUE, "village_class.update", _
                                        "����� ���� ��� ���� � �����������"
        Else
            Err.Raise Err.Number, "street_class.update", Err.Description
        End If
    End If
        
    Call reloadVillageList
End Sub


Private Sub reloadVillageList()
' ----------------------------------------------------------------------------
' ���������� ������ ��
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    address_village_list.reload
End Sub


Private Sub flushvalues()
'-----------------------------------------------------------------------------
' ��������� ������� ��������
' Last update: 22.03.2018
'-----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
