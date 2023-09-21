VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkKindForm 
   Caption         =   "���������� ����� �����"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "WorkKindForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkKindForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New workkind_class


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ��������� �����, ���������� �����
' Last update: 18.05.2018
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
        For i = 1 To FormWorkKindEnum.fwMax
            .add
        Next i
        .Item(FormWorkKindEnum.fwId + 1).text = "���"
        .Item(FormWorkKindEnum.fwName + 1).text = "��������"
        .Item(FormWorkKindEnum.fwWorkType + 1).text = "��� �����"
    End With
    
    Call reloadListView
    Call reloadComboBox(rcmWorkType, Me.ComboBoxWT, addAllItems:=True)
    Call reloadComboBox(rcmWorkType, Me.ComboBoxWTChange)
    
    Me.TextBoxName.SetFocus
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��������� ������ �������� � ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New workkind_class
    curItem.initial CLng(Item)
    Me.TextBoxName = curItem.Name
    Call selectComboBoxValue(Me.ComboBoxWTChange, curItem.workType.Id)
    Me.LabelCurItem.Caption = curItem.Name & " (" & curItem.workType.Name & ")"
End Sub


Private Sub ListViewList_ColumnClick( _
                                ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' ���������� ��� ������ �� �������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub ButtonClose_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� �����
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ����������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New workkind_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub ButtonChange_Click()
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


Private Sub ButtonDelete_Click()
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


Private Sub ComboBoxWT_Change()
' ----------------------------------------------------------------------------
' ��� ������ ���� ����� ������������ ������ �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call reloadListView
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' ���������� ������� ListView
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim listX As ListItem
    Dim i As Long, j As Long
    Dim curList As workkind_list
    Dim wtId As Long
    
    ' ���������� �������
    wtId = IIf(Me.ComboBoxWT.ListIndex = -1, NOTVALUE, Me.ComboBoxWT.Value)
    Set curList = New workkind_list
    curList.initial wtId
    Me.ListViewList.ListItems.Clear
    For i = 1 To curList.count
        Set curItem = curList(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormWorkKindEnum.fwMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormWorkKindEnum.fwName).text = curItem.Name
        listX.ListSubItems(FormWorkKindEnum.fwWorkType).text = curItem.workType.Name
        Set curItem = Nothing
    Next i
    Set listX = Nothing

    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' ����������/��������� ���� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If formNotFill Then
        MsgBox "��������� �� ��� ����!", vbInformation + vbOKOnly, "������"
                                                            
        GoTo cleanHandler
    End If
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
                        newWT:=Me.ComboBoxWTChange.Value, _
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
    formNotFill = (StrComp(Trim(Me.TextBoxName.Value), "") = 0 Or _
                    Me.ComboBoxWTChange.ListIndex = -1)
End Function


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' ������� ���� ��������� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.ComboBoxWTChange.ListIndex = -1
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub
