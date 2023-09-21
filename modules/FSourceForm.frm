VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSourceForm 
   Caption         =   "��������� ��������������"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "FSourceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FSourceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As fsource


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� ����� - ���������� ����� ���� ���������
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    If curItem Is Nothing Then
        Me.Caption = "���������� ��������� ��������������"
        Me.ButtonSave.Caption = "��������"
    Else
        Me.Caption = "��������� ��������� ��������������"
        Me.TextBoxName = curItem.Name
        Me.TextBoxNote = curItem.Note
        Me.CheckBoxFromSA = curItem.FromSubaccount
        Me.ButtonSave.Caption = "���������"
    End If
    Me.Caption = Me.Caption & ". ������ " & AppConfig.DBServer
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ����������
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    If formFill Then
        If curItem Is Nothing Then
            Set curItem = New fsource
            curItem.create pName:=Trim(Me.TextBoxName), _
                            pNote:=Trim(Me.TextBoxNote), _
                            pFromSubaccount:=Me.CheckBoxFromSA
        Else
            curItem.update newName:=Trim(Me.TextBoxName), _
                            newNote:=Trim(Me.TextBoxNote), _
                            newFromSubaccount:=Me.CheckBoxFromSA
        End If
        Set curItem = Nothing
        Unload Me
    Else
        MsgBox "��������� �� ��� ����"
    End If
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ������
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    Set curItem = Nothing
    Unload Me
End Sub

Private Function formFill() As Boolean
' ----------------------------------------------------------------------------
' �������� ���������� �����
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    formFill = Len(Trim(Me.TextBoxName)) > 0
End Function

