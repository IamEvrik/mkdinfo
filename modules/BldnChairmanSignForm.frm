VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BldnChairmanSignForm 
   Caption         =   "���������� ������� ����������"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
   OleObjectBlob   =   "BldnChairmanSignForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BldnChairmanSignForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As base_form_class

' ����� ���������� �������


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ���������� ����� ��� ��������
' 19.10.2022
' ----------------------------------------------------------------------------
    If curItem Is Nothing Then Unload Me
    Call reloadComboBox(rcmTermDESC, Me.ComboBoxTerm, defValue:=terms.LastTerm.Id)
    
End Sub


Private Sub BtnGetSignFile_Click()
' ----------------------------------------------------------------------------
' ����� ����� �������
' 16.10.2022
' ----------------------------------------------------------------------------
    Dim imageName As String
    imageName = Application.GetOpenFilename("�������� (*.png;*.jpg),*.png;*.jpg")
    If imageName <> "False" Then
        Me.SignFileName = imageName
    End If
End Sub


Private Sub CheckBoxNotChairman_Change()
' ----------------------------------------------------------------------------
' ������� ����������� ����������?
' 16.10.2022
' ----------------------------------------------------------------------------
    If Me.CheckBoxNotChairman.Value Then
        Me.TextBoxSignOwner.Enabled = True
    Else
        Me.TextBoxSignOwner.Enabled = False
    End If
End Sub


Private Sub CommandButtonCancel_Click()
' ----------------------------------------------------------------------------
' ����� �� ����� ��� ����������
' 16.10.2022
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Property Get isFormFill() As Boolean
' ----------------------------------------------------------------------------
' �������� ���������� �����
' 20.10.2022
' ----------------------------------------------------------------------------
' 20.10.2022 ����� �������� �� ���� �������, �.�. ����� ������ �������,
' ��� �����������, ��� ���������� ����� �������
'    If Me.SignFileName.Caption = "" Then
'        isFormFill = False
'        Exit Sub
'    End If
    
    If Not Me.CheckBoxNotChairman.Value Then
        isFormFill = True
        Exit Sub
    End If
    
    isFormFill = Me.TextBoxSignOwner.Value <> ""
    
    isFormFill = Not (Me.CheckBoxNotChairman.Value And Me.TextBoxSignOwner.Value = "")
End Property


Private Function isAttachedSign() As Boolean
' ----------------------------------------------------------------------------
' ������� �� �������
' 20.10.2022
' ----------------------------------------------------------------------------
    isAttachedSign = (Me.SignFileName.Caption <> "")
End Function


Private Sub CommandButtonAdd_Click()
' ----------------------------------------------------------------------------
' ���������� �������
' 20.10.2022
' ----------------------------------------------------------------------------

    On Error GoTo errHandler
    
    If Not Me.isFormFill() Then
        MsgBox "�� ��������� ��� ����������� ����", vbInformation, _
                "��������� ����������"
        Exit Sub
    End If
    
    If Not isAttachedSign Then
        Dim atAns As Integer
        atAns = MsgBox("�� ������ ���� �������, �� ������ ����������", _
                vbYesNo, "�� ������� �������")
        If atAns = vbNo Then Exit Sub
    End If
    
    If Not Me.CheckBoxNotChairman Then
        Me.TextBoxSignOwner.Value = ""
    End If
    curItem.update Me, False

errHandler:
    If errorHasNoPrivilegies(Err.Description) Then
        MsgBox "�� ������� ����", vbExclamation, "������ ����������"
    ElseIf errorHasNoValues(Err.Number) Then
        MsgBox "������ �� �����", vbExclamation, "������ ����������"
    ElseIf errorNotUnique(Err.Description) Then
        MsgBox "������� �� ���� ����� ��� ���� � ����", vbExclamation, "������ ����������"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & vbCr & Err.Source & vbCr & Err.Description, _
                                                vbCritical, "������ ����������"
    End If
End Sub

