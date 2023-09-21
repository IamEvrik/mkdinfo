VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GlobalWorkForm 
   Caption         =   "���� ��������"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "GlobalWorkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GlobalWorkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New globalWorkType_class


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    globalWorkType_list.reload
    Call reloadListView
End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ��������� �����, ���������� �����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    ' � ���� �������
    Me.ListViewList.View = lvwReport
    ' ���������� ��� ������
    Me.ListViewList.FullRowSelect = True
    ' ������ ��������� �������� � ����� ListView
    Me.ListViewList.LabelEdit = lvwManual
    ' ��������� ��������
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To FormGWTEnum.fgwtMax
            .add
        Next i
        .Item(FormGWTEnum.fgwtId + 1).text = "���"
        .Item(FormGWTEnum.fgwtName + 1).text = "��������"
        .Item(FormGWTEnum.fgwtNote + 1).text = "����������"
    End With
    
    Call reloadListView
    
    Me.TextBoxName.SetFocus
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��������� ������ �������� � ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = globalWorkType_list(CStr(Item))
    Me.TextBoxName = curItem.Name
    Me.TextBoxNote = curItem.Note
    Me.LabelCurItem.Caption = curItem.Name
End Sub


Private Sub ListViewList_ColumnClick( _
                                ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' ���������� ��� ������ �� �������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� �����
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ����������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New globalWorkType_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub BtnChange_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ���������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ������� ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub BtnDelete_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ��������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        ' ������ �������������
        If Not ConfirmDeletion(curItem.Name) Then Exit Sub
        
        On Error GoTo errHandler
        curItem.delete
        
        ' �����������
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler
        
errHandler:
        If Err.Number = ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "������ ��������"
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
cleanHandler:
    End If
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' ���������� ������� ListView
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim listX As ListItem
    Dim i As Long, j As Long
    
    ' ���������� �������
    Me.ListViewList.ListItems.Clear
    For i = 1 To globalWorkType_list.count
        Set curItem = globalWorkType_list(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormGWTEnum.fgwtMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormGWTEnum.fgwtName).text = curItem.Name
        listX.ListSubItems(FormGWTEnum.fgwtNote).text = curItem.Note
        Set curItem = Nothing
    Next i
    Set listX = Nothing

    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' ����������/��������� ���� �������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If formNotFill Then
        MsgBox "��������� �� ��� ����!", vbInformation + vbOKOnly, "������"
                                                            
        GoTo cleanHandler
    End If
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
                        newNote:=Me.TextBoxNote.Value, _
                        addNew:=addFlag
        
        ' ����������� �����
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler
        
errHandler:
        If Err.Number = ERROR_NOT_UNIQUE Then
            MsgBox Err.Description, vbInformation, "������"
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
        
cleanHandler:
    End If
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    formNotFill = (StrComp(Trim(Me.TextBoxName.Value), "") = 0)
End Function


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' ������� ���� ��������� �����
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.TextBoxNote.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub
