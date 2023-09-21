VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UKForm 
   Caption         =   "���������� ����������� ��������"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "UKForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UKForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New uk_class


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    uk_list.reload
    Call reloadListView
End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ��������� �����, ���������� �����
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    Me.Caption = Me.Caption & " " & AppConfig.DBServer
    
    ' � ���� �������
    Me.ListViewList.View = lvwReport
    ' ���������� ��� ������
    Me.ListViewList.FullRowSelect = True
    ' ������ ��������� �������� � ����� ListView
    Me.ListViewList.LabelEdit = lvwManual
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To FormMCEnum.fmcMax
            .add
        Next i
        .Item(FormMCEnum.fmcID + 1).text = "���"
        .Item(FormMCEnum.fmcName + 1).text = "��������"
        .Item(FormMCEnum.fmcReportName + 1).text = "����. ��������"
        .Item(FormMCEnum.fmcNotManage + 1).text = "�����������"
        .Item(FormMCEnum.fmcChiefEngineer + 1).text = "������� �������"
        .Item(FormMCEnum.fmcDirector + 1).text = "��������"
    End With
    Me.ListViewList.Sorted = True
    Me.ListViewList.SortKey = FormMCEnum.fmcID
    
    Call reloadListView
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� ����� - ������� �����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��������� ������ ����������� �������� � ������
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set curItem = uk_list(CStr(Item))
    ' ���������� �����
    Me.TextBoxName = curItem.Name
    Me.TextBoxReportName = curItem.reportName
    Me.CheckBoxNotManage = curItem.notManage
    Me.LabelCurItem.Caption = curItem.Name
End Sub


Private Sub ListViewList_ColumnClick(ByVal ColumnHeader As _
                                                    MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' ���������� ��� ������ �� �������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� �����
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ������� �����
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ���������� ��
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New uk_class
    Call process(addFlag:=True)
End Sub


Private Sub BtnChange_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ��������� ��
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub BtnDelete_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� ��
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        ' ������������� ��������
        If Not ConfirmDeletion(curItem.Name) Then Exit Sub
        
        On Error GoTo errHandler
        curItem.delete
        
        Call reloadListView
        Call clearTextBox
        
errHandler:
        If Err.Number = ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "������"
        Else
            MsgBox Err.Number & "-->" & Err.Description, vbCritical
        End If
    End If
End Sub


Private Sub BtnShowEmployee_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "����������", ����� ����� �����������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        Me.Hide
        Call RunEmployeeForm(curItem)
    End If
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' ���������� ������� ListView
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As ListItem
    
    ' ���������� �������
    Me.ListViewList.ListItems.Clear
    For i = 1 To uk_list.count
        Set curItem = uk_list(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormMCEnum.fmcMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormMCEnum.fmcName).text = curItem.Name
        listX.ListSubItems(FormMCEnum.fmcReportName).text = curItem.reportName
        listX.ListSubItems(FormMCEnum.fmcDirector).text = curItem.Director.FIO
        listX.ListSubItems(FormMCEnum.fmcChiefEngineer).text = _
                                                    curItem.ChiefEngineer.FIO
        listX.ListSubItems(FormMCEnum.fmcNotManage).text = _
                                            BoolToYesNo(Not curItem.notManage)
    Next i
    Set curItem = Nothing
    
    ' �������������� ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set listX = Nothing
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' ����������/��������� � ����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Or addFlag Then
        If formNotFill Then
            MsgBox "��������� �� ��� ����������� ����", vbInformation, _
                                                                    "������"
            Exit Sub
        End If
            
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
                        newReport:=Me.TextBoxReportName.Value, _
                        newNotManage:=Me.CheckBoxNotManage.Value, _
                        addNew:=addFlag
                
        Call reloadListView
        Call clearTextBox
    End If
    GoTo cleanHandler

errHandler:
    If Err.Number = ERROR_NOT_UNIQUE Then
        MsgBox Err.Description, vbInformation, "������"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

cleanHandler:
End Sub


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' ������� ���� ��������� �����
' Last update: 04.03.2016
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.TextBoxReportName.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
    Me.CheckBoxNotManage.Value = False
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������� �����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    formNotFill = False
    If StrComp(Trim(Me.TextBoxName.Value), "") = 0 Or _
                    StrComp(Trim(Me.TextBoxReportName.Value), "") = 0 Then
        formNotFill = False
    End If
End Function
