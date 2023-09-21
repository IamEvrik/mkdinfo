VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleSelectionForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "SimpleSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fCurItem As Long                ' ������� ��������� ��������
Private f_CurText As String             ' ����� ���������� ��������


Property Get selectedItem() As Long
' ----------------------------------------------------------------------------
' ������� ���������� ��������
' Last update: 10.04.2019
' ----------------------------------------------------------------------------
    selectedItem = fCurItem
End Property

Property Get selectedText() As String
' ----------------------------------------------------------------------------
' ������� ���������� ������
' 09.09.2021
' ----------------------------------------------------------------------------
    selectedText = f_CurText
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' ������ ��������� ����� �� �������
' Last update: 10.04.2019
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


Private Sub BtnChoose_Click()
' ----------------------------------------------------------------------------
' ��� ������� �� ������ ���������� �������� ��������
' 09.09.2021
' ----------------------------------------------------------------------------
    fCurItem = Me.ComboBox1.Value
    f_CurText = Me.ComboBox1.text
    Me.Hide
End Sub
