VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TitleForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895
   OleObjectBlob   =   "TitleForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TitleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ����� ������ ���������
' Author: Evrik
' ----------------------------------------------------------------------------

Private f_titles As Collection
Private Const title_folder As String = "config"


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����, �������� ����������
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    Me.Caption = "����� ���������. ������: " & AppConfig.DBServer
    Call loadTitles
    Me.LabelDescription.Caption = "��� �������� �� ��������� ������ " & _
                                    "����������� Ctrl+Enter"
End Sub


Private Sub ComboBoxTitles_Change()
' ----------------------------------------------------------------------------
' ���������� ���������� ���� ��� ��������� ������
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxTitles.ListIndex > -1 Then
        Me.TextBoxTitle.Value = f_titles(Me.ComboBoxTitles.Value)
    End If
End Sub


Private Sub ButtonOk_Click()
' ----------------------------------------------------------------------------
' ���������� ��������� � �����
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxTitles.ListIndex > -1 Then
        If StrComp(Me.TextBoxTitle, f_titles(Me.ComboBoxTitles.Value), _
                                                    vbBinaryCompare) <> 0 Then
            Call setFileText(getFullFileName(Me.ComboBoxTitles.Value), _
                                                        Me.TextBoxTitle.Value)
        End If
        If Me.TextBoxTitle.Value = "" Then Me.TextBoxTitle.Value = NOTSTRING
        Me.Hide
    Else
        MsgBox "�������� ����� ��� ��������� ����� ������"
    End If
End Sub


Private Sub ButtonSaveAs_Click()
' ----------------------------------------------------------------------------
' ���������� ������ �������
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    With SimpleAddForm
        .Caption = "������� �������� �������"
        .onlyText = True
        .Show
        If .TextBoxName.Value <> "" Then
            Call setFileText(getFullFileName(.TextBoxName.Value), _
                                                        Me.TextBoxTitle.Value)
            Call loadTitles
        End If
    End With
End Sub


Private Sub ButtonDelete_Click()
' ----------------------------------------------------------------------------
' �������� �������
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxTitles.ListIndex > -1 Then
        If ConfirmDeletion("������ " & Me.ComboBoxTitles) Then
            Kill getFullFileName(Me.ComboBoxTitles.Value)
            Call loadTitles
        End If
    End If
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' ������ ������ (����� ������������� � NOTSTRING)
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    Me.TextBoxTitle.Value = NOTSTRING
    Me.Hide
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' ��� �������� �� ������� ���������� NOTSTRING
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    If CloseMode = 0 Then
        MsgBox "������� ���� �� ������"
        Cancel = True
    End If
End Sub


Sub loadTitles()
' ----------------------------------------------------------------------------
' ��������� ������ ����������
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    Dim fName As String, fContent As String, contentKey As String
    Dim dirName As String
    
    Set f_titles = New Collection
    Me.ComboBoxTitles.Clear
    Me.TextBoxTitle.Value = ""
    dirName = getThisPath & title_folder & Application.PathSeparator
    fName = Dir(dirName & "report_title *.txt")
    Do While fName <> ""
        contentKey = getSpecify(fName)
        fContent = getFileText(dirName & fName)
        f_titles.add Item:=fContent, Key:=contentKey
        Me.ComboBoxTitles.AddItem contentKey
        fName = Dir
    Loop
    If Me.ComboBoxTitles.ListCount > 0 Then Me.ComboBoxTitles.ListIndex = 0
End Sub


Function getFullFileName(specify As String) As String
' ----------------------------------------------------------------------------
' ������ ���� � ����� �� ��� ������������
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    getFullFileName = getThisPath & title_folder & _
                Application.PathSeparator & "report_title " & specify & ".txt"
End Function


Function getSpecify(fileName As String) As String
' ----------------------------------------------------------------------------
' ������������ �� ������� ����� �����
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    getSpecify = Replace(fileName, getThisPath & title_folder & _
                                                Application.PathSeparator, "")
    getSpecify = Replace(getSpecify, "report_title ", "")
    getSpecify = Replace(getSpecify, ".txt", "")
End Function
