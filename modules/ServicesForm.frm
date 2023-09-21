VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServicesForm 
   Caption         =   "������ �����"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13350
   OleObjectBlob   =   "ServicesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ServicesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As Service


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� �����, �������� ���� � ���������
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & " " & AppConfig.DBServer
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����, ���������� ������ �����
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Call ReloadServices
End Sub


Private Sub ListBoxServices_Click()
' ----------------------------------------------------------------------------
' ��� ������ ������ ���������� ������ �������
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    If Me.ListBoxServices.ListIndex > -1 Then
        Set curItem = services(CStr(Me.ListBoxServices.Value))
        Call ReloadServiceModes
    End If
End Sub


Private Sub ButtonAddService_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ���������� ������
' 28.06.2022
' ----------------------------------------------------------------------------
    Set curItem = New Service
    curItem.showForm False
    Call ReloadServices
End Sub


Private Sub ListBoxServices_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' ----------------------------------------------------------------------------
' ��� ������� ������ ����� ��������� �������� ������
' 28.06.2022
' ----------------------------------------------------------------------------
    Call ChangeService
End Sub


Private Sub ButtonChangeService_Click()
' ----------------------------------------------------------------------------
' ��������� �������� ������
' 28.06.2022
' ----------------------------------------------------------------------------
    Call ChangeService
End Sub


Private Sub ChangeService()
' ----------------------------------------------------------------------------
' ��������� ������
' 28.06.2022
' ----------------------------------------------------------------------------
    If Not curItem Is Nothing Then
        curItem.showForm isChange:=True
        Call ReloadServices
    End If

End Sub


Private Sub ButtonDeleteService_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� ������
' Last update: 22.04.2019
' ----------------------------------------------------------------------------
    If Not curItem Is Nothing Then
        On Error GoTo errHandler
        If ConfirmDeletion(curItem.Name) Then
            curItem.delete
        End If
        Call ReloadServices
errHandler:
        Set curItem = Nothing
        If Err.Number <> 0 Then
            MsgBox Err.Description & vbCr & Err.Source
            Err.Clear
        End If
    End If
End Sub


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' ���������� ������ ����� (�.�. ����������� �����)
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    services.reload
    Call ReloadServices
End Sub


Private Sub ButtonAddMode_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ���������� ������ �����������
' 29.06.2022
' ----------------------------------------------------------------------------
    If Not curItem Is Nothing Then
        Dim tmpMode As New service_mode
        tmpMode.addEmpty curItem.Id
        tmpMode.showForm isChange:=False
        Set tmpMode = Nothing
        Call ReloadServices(Me.ListBoxServices)
    End If
End Sub


Private Sub ButtonDeleteMode_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� ������
' Last update: 22.04.2019
' ----------------------------------------------------------------------------
    If Me.ListBoxServiceModes.ListIndex > -1 Then
        On Error GoTo errHandler
        Dim ctmp As New service_mode
        ctmp.add Id:=Me.ListBoxServiceModes.Value, _
                serviceId:=curItem.Id, _
                Name:=Me.ListBoxServiceModes.text
        If ConfirmDeletion(curItem.Name & " " & ctmp.Name) Then
            ctmp.delete
        End If
        Call ReloadServices(Me.ListBoxServices)
errHandler:
        Set ctmp = Nothing
        If Err.Number <> 0 Then
            MsgBox Err.Description & vbCr & Err.Source
            Err.Clear
        End If
    End If
End Sub


Private Sub ListBoxServiceModes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' ----------------------------------------------------------------------------
' ��� ������� ������ ����� ��������� �������� ������
' Last update: 16.10.2018
' ----------------------------------------------------------------------------
    Call ChangeMode
End Sub


Private Sub ChangeMode()
' ----------------------------------------------------------------------------
' ��������� �������� ������
' 29.06.2022
' ----------------------------------------------------------------------------
    If Not curItem Is Nothing And Me.ListBoxServiceModes.ListIndex > -1 Then
        Dim ctmp As New service_mode
        ctmp.add Id:=Me.ListBoxServiceModes.Value, _
                serviceId:=curItem.Id, _
                Name:=Me.ListBoxServiceModes.text
        ctmp.showForm isChange:=True
        Set ctmp = Nothing
        Call ReloadServices
    End If
End Sub


Private Sub ReloadServices(Optional curService As Long = NOTVALUE)
' ----------------------------------------------------------------------------
' ���������� ������ �����
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Me.ListBoxServiceModes.Clear
    Call reloadComboBox(rcmServices, Me.ListBoxServices)
    If Not curItem Is Nothing Then
        Call selectComboBoxValue(Me.ListBoxServices, curItem.Id)
    End If
End Sub


Private Sub ReloadServiceModes()
' ----------------------------------------------------------------------------
' ���������� ������ ������� �����
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Call reloadComboBox(rcmServiceModes, Me.ListBoxServiceModes, _
                                                    CLng(Me.ListBoxServices))
End Sub
