VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VillageForm 
   Caption         =   "������ ��������� �������"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13755
   OleObjectBlob   =   "VillageForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VillageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curItem As New address_village_class


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    address_village_list.reload
    Call reloadListView
End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� ����� - ���������� ������ �� � ����������� ������ ��
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    ' ������� ��������� ���������� ListView
    ' � ���� �������
    Me.ListViewList.View = lvwReport
    ' ���������� ��� ������
    Me.ListViewList.FullRowSelect = True
    ' ������ ��������� �������� � ����� ListView
    Me.ListViewList.LabelEdit = lvwManual
    ' ��������� ��������
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To FormVillageEnum.fvMax
            .add
        Next i
        .Item(FormVillageEnum.fvId + 1).text = "���"
        .Item(FormVillageEnum.fvMD + 1).text = "������������� �����������"
        .Item(FormVillageEnum.fvName + 1).text = "��������"
        .Item(FormVillageEnum.fvSite + 1).text = "��� �����"
    End With

    address_village_list.reload
    Call reloadListView
    Call reloadComboBox(rcmMd, Me.ComboBoxMD)
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' �������� �����
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Set curItem = Nothing
End Sub


Private Sub ListViewList_ColumnClick( _
                            ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' ��� ������ �� ������� - ���������� �� ����
' Last update: 12.02.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��� ������ ������ ������ ����������� ��������� ����
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Set curItem = address_village_list(CStr(Item))
    
    Me.TextBoxName.Value = curItem.Name
    Me.TextBoxSite.Value = curItem.Site
    Call selectComboBoxValue(Me.ComboBoxMD, curItem.Municipal_district.Id)
    Me.LabelCurItem.Caption = curItem.Name & " (" & _
                                        curItem.Municipal_district.Name & ")"
End Sub


Private Sub CloseButton_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "�������"
' Last update: 24.02.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "�������� �����"
' Last update: 23.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub AddButton_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "��������"
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New address_village_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub ChangeButton_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "��������"
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub DeleteButton_Click()
' -----------------------------------------------------------------------------
' About: ��������� ������� ������ "�������"
' Last update: 25.03.2018
' -----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        ' ������������� ��������
        If Not ConfirmDeletion(curItem.Name & " " & _
                                curItem.Municipal_district.Name) Then Exit Sub
        On Error GoTo errHandler
        curItem.delete
        
        ' ����������� �����
        Call reloadListView
        Call clearTextBox

errHandler:
        If ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "������"
        Else
            MsgBox Err.Number & "-->" & Err.Description, vbCritical
        End If
    End If
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' ���������� ������ ����
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As ListItem
    
    ' ���������� �������
    Me.ListViewList.ListItems.Clear
    For i = 1 To address_village_list.count
        Set curItem = address_village_list(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormVillageEnum.fvMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormVillageEnum.fvMD).text = _
                                            curItem.Municipal_district.Name
        listX.ListSubItems(FormVillageEnum.fvName).text = curItem.Name
        listX.ListSubItems(FormVillageEnum.fvSite).text = curItem.Site
    Next i
    
    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set listX = Nothing
    Set curItem = Nothing
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' ����������/��������� � ����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        If formNotFill Then
            MsgBox "��������� �� ��� ����", vbInformation + vbOKOnly, "������"
        Else
            On Error GoTo errHandler
            curItem.update newName:=Me.TextBoxName.Value, _
                        newMd:=Me.ComboBoxMD.Value, _
                        newSite:=Me.TextBoxSite.Value, _
                        addNew:=addFlag
            
            ' ���������� ������
            Call reloadListView
            Call clearTextBox
        End If
    End If
    GoTo cleanHandler

errHandler:
    If ERROR_NOT_UNIQUE Then
        MsgBox Err.Description, vbInformation, "������"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

cleanHandler:
End Sub


Private Sub clearTextBox()
' -----------------------------------------------------------------------------
' ������� ���� �����
' Last update: 23.03.2018
' -----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.TextBoxSite.Value = ""
    Me.ComboBoxMD.ListIndex = -1
    Set curItem = Nothing
    Me.LabelCurItem.Caption = "�� �������"
End Sub


Private Function formNotFill() As Boolean
' -----------------------------------------------------------------------------
' ����������� ���������� ����������� �����
' Last update: 22.03.2018
' -----------------------------------------------------------------------------
    formNotFill = (StrComp(Trim(Me.TextBoxName.Value), "") = 0 Or _
                                                Me.ComboBoxMD.ListIndex = -1)
End Function
