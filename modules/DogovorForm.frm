VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DogovorForm 
   Caption         =   "���� ���������"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "DogovorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DogovorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New dogovor_class


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
        For i = 1 To FormImprovementEnum.fiMax
            .add
        Next i
        .Item(FormImprovementEnum.fiId + 1).text = "���"
        .Item(FormImprovementEnum.fiName + 1).text = "��������"
        .Item(FormImprovementEnum.fiShortName + 1).text = "����������� ��������"
    End With
    
    Call reloadListView
    
    Me.TextBoxFullName.SetFocus
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��������� ������ �������� � ������
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Set curItem = dogovor_list(CStr(Item))
    Me.TextBoxFullName = curItem.FullName
    Me.TextBoxShortName = curItem.ShortName
    Me.LabelCurItem.Caption = curItem.FullName
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
    Set curItem = New dogovor_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    dogovor_list.reload
    Call reloadListView
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
        If Not ConfirmDeletion(curItem.FullName) Then Exit Sub
        
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
    For i = 1 To dogovor_list.count
        Set curItem = dogovor_list(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormImprovementEnum.fiMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormImprovementEnum.fiName).text = curItem.FullName
        listX.ListSubItems(FormImprovementEnum.fiShortName).text = curItem.ShortName
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
        curItem.update newFullName:=Me.TextBoxFullName.Value, _
                        newShortName:=Me.TextBoxShortName, _
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
    formNotFill = (StrComp(Trim(Me.TextBoxFullName.Value), "") = 0 Or _
                    StrComp(Trim(Me.TextBoxShortName.Value), "") = 0)
End Function


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' ������� ���� ��������� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Me.TextBoxFullName.Value = ""
    Me.TextBoxShortName.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub
