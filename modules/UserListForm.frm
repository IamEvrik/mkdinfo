VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserListForm 
   Caption         =   "������ �������������"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "UserListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� �����
' Last update: 17.09.2018
' ----------------------------------------------------------------------------
    Dim i As Long

    Me.Caption = Me.Caption & " " & AppConfig.DBServer
    
    With Me.ListViewList
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormUserList.fulMax
                .add
            Next i
            .Item(FormUserList.fulId + 1).text = "���"
            .Item(FormUserList.fulFIO + 1).text = "���"
            .Item(FormUserList.fulIsActive + 1).text = "�������"
            .Item(FormUserList.fulLogin + 1).text = "�����"
        End With
    End With
    Call reloadList
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� ����� - ����� ���������
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Me.ListViewList.selectedItem.Selected = False
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' ���������� ������������
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    With UserChangeForm
        .formMode = ufmAdd
        .Show vbModal
    End With
End Sub


Private Sub ButtonChangeName_Click()
' ----------------------------------------------------------------------------
' ��������� ���
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To Me.ListViewList.ListItems.count
        If Me.ListViewList.ListItems(i).Selected Then
            With UserChangeForm
                .formMode = ufmChangeName
                .userId = Me.ListViewList.selectedItem
                .TextBoxFio.Value = Me.ListViewList.selectedItem.ListSubItems( _
                                                    FormUserList.fulFIO).text
                .TextBoxLogin.Value = Me.ListViewList.selectedItem.ListSubItems( _
                                                    FormUserList.fulLogin).text
                .Show
            End With
        End If
    Next i
End Sub


Private Sub ButtonChangePwd_Click()
' ----------------------------------------------------------------------------
' ��������� ������
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To Me.ListViewList.ListItems.count
        If Me.ListViewList.ListItems(i).Selected Then
            With UserChangeForm
                .formMode = ufmChangePassword
                .userId = Me.ListViewList.selectedItem
                .TextBoxLogin.Value = Me.ListViewList.selectedItem.ListSubItems( _
                                                    FormUserList.fulLogin).text
                .Show
            End With
        End If
    Next i
End Sub


Private Sub ButtonUnactive_Click()
' ----------------------------------------------------------------------------
' ���������� ������������
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim ans As Long
    For i = 1 To Me.ListViewList.ListItems.count
        If Me.ListViewList.ListItems(i).Selected Then
            ans = MsgBox("�� ������������� ������ ������������� " & _
                    "������������ " & _
                    Me.ListViewList.ListItems(i).ListSubItems( _
                    FormUserList.fulLogin).text & "?", vbYesNo + vbQuestion, _
                    "����������� ����������")
            If ans = vbNo Then Exit Sub
            Dim cmd As ADODB.Command
            
            Set cmd = New ADODB.Command
            Set cmd.ActiveConnection = DBConnection.Connection
            cmd.CommandText = "adm_block_user"
            cmd.CommandType = adCmdStoredProc
            cmd.NamedParameters = True
            cmd.Parameters("itemId").Value = Me.ListViewList.selectedItem
            cmd.Parameters("userId").Value = CurrentUser.userId
            cmd.Execute
            Set cmd = Nothing
        End If
    Next i
    Call reloadList
End Sub


Public Sub reloadList()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As ListItem
    Dim tmpList As users
    Dim curItem As user
    
    ' ���������� �������
    Set tmpList = New users
    Me.ListViewList.ListItems.Clear
    For i = 1 To tmpList.count
        Set curItem = tmpList(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormUserList.fulMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormUserList.fulFIO).text = _
                                            curItem.Name
        listX.ListSubItems(FormUserList.fulIsActive).text = _
                                            BoolToYesNo(curItem.is_active)
        listX.ListSubItems(FormUserList.fulLogin).text = curItem.login
    Next i
    ' ������� �������������� ��������� ������ ������
    Me.ListViewList.selectedItem.Selected = False
    
    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set listX = Nothing
    Set curItem = Nothing
End Sub
