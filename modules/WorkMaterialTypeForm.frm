VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkMaterialTypeForm 
   Caption         =   "���������� ��������� �����"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   OleObjectBlob   =   "WorkMaterialTypeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkMaterialTypeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As material_type


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� ����� - ���������� ����� ���� ���������
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    If curItem Is Nothing Then
        Me.Caption = "���������� ��������� �����"
        Me.BtnSave.Caption = "��������"
    Else
        Me.Caption = "��������� ��������� �����"
        Me.TextBoxName = curItem.Name
        Me.CheckBoxIsTransport = curItem.IsTransport
        Me.BtnSave.Caption = "���������"
    End If
    Me.Caption = Me.Caption & ". " & AppConfig.DBServer
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' ����������� �����
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    Set curItem = Nothing
End Sub


Private Sub BtnSave_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ����������
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    Call process
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' �������� ����� �� ������ ������
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub process()
' ----------------------------------------------------------------------------
' ����������/���������
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    Dim addFlag As Boolean
    
    On Error GoTo errHandler

    If formNotFill Then
        MsgBox "��������� �� ��� ����!", vbInformation + vbOKOnly, "������"
        GoTo cleanHandler
    End If
    
    addFlag = False
    If curItem Is Nothing Then
        Set curItem = New material_type
        addFlag = True
    End If
    
    curItem.update newName:=Me.TextBoxName.Value, _
                    newIsTransport:=Me.CheckBoxIsTransport, _
                    addNew:=addFlag
    Unload Me
    GoTo cleanHandler
    
errHandler:
    Dim errMsg As String
    If errorHasNoPrivilegies(Err.Description) Then
        errMsg = "�� ������� ����"
    ElseIf errorNotUnique(Err.Description) Then
        errMsg = "������ �������� ��� ����������"
    ElseIf errorStopDelete(Err.Description) Then
        errMsg = "������ �������� ��������� ���"
    Else
        errMsg = Err.Description
    End If
    MsgBox errMsg, vbCritical, "������"
    
cleanHandler:
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������� ����������� �����
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    formNotFill = (Len(Trim(Me.TextBoxName.Value)) = 0)
End Function
