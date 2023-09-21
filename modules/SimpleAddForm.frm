VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleAddForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "SimpleAddForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleAddForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As Object        ' ������� ������
Public isChange As Boolean      ' ������� ���������/����������
Public parentId As Long         ' ��� ��������
Public onlyText As Boolean


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & " ������ " & AppConfig.DBServer
    If isChange Then Me.TextBoxName.Value = curItem.Name
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' ������ ������, �������� �����
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' ���������� ����������
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    If Not formFill Then
        MsgBox "��������� ����", vbExclamation
        Me.TextBoxName.SetFocus
        Exit Sub
    End If
    If onlyText Then
        Me.Hide
    ElseIf Me.parentId > 0 Then
        curItem.update parentId, Me.TextBoxName, Not isChange
    Else
        curItem.update newName:=Me.TextBoxName, addNew:=Not (isChange)
    End If
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


Private Function formFill() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������� �����
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    If Len(Trim(Me.TextBoxName.Value)) = 0 Then
        formFill = False
    Else
        formFill = True
    End If
End Function
