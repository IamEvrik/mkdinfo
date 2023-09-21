VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmployeeForm 
   Caption         =   "������ �����������"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13650
   OleObjectBlob   =   "EmployeeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmployeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private posId As PositionStatusEnum
Private curItem As employee_class
Public curMC As uk_class


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� �����
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    Me.Caption = Me.Caption & " " & AppConfig.DBServer
    
    Me.ListViewList.View = lvwReport     ' � ���� �������
    Me.ListViewList.FullRowSelect = True ' ���������� ��� ������
    ' ������ ��������� �������� � ����� ListView
    Me.ListViewList.LabelEdit = lvwManual
    ' ��������� ��������
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To FormEmployeeEnum.feMax
            .add
        Next i
        .Item(FormEmployeeEnum.feId + 1).text = "���"
        .Item(FormEmployeeEnum.feFirstName + 1).text = "���"
        .Item(FormEmployeeEnum.feLastName + 1).text = "�������"
        .Item(FormEmployeeEnum.fePosition + 1).text = "���������"
        .Item(FormEmployeeEnum.feSecondName + 1).text = "��������"
        .Item(FormEmployeeEnum.feSignReport + 1).text = "������� ������"
    End With
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' ����������� �����
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Set curItem = Nothing
    Set curMC = Nothing
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����, ���������� �����
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    ' ����� ������ ��������� ��������������, ������ �� ����� ��
    If curMC Is Nothing Then
        Unload Me
    ElseIf curMC.Id = NOTVALUE Then
        Unload Me
    Else
        Me.LabelOrgName.Caption = curMC.Name
        Call reloadListView
        Call clearTextBox
    End If
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��� ������ �� ���������� - ���������� ��������� �����
' Last update: 28.03.2017
' ----------------------------------------------------------------------------
    Set curItem = curMC.employees(CStr(Item))
    Me.TBFirstName.Value = curItem.FirstName
    Me.TBLastName.Value = curItem.LastName
    Me.TBPosition.Value = curItem.Position
    Me.TBSecondName.Value = curItem.SecondName
    Me.TBSign.Value = curItem.ReportSign
    Me.LabelCurItem.Caption = curItem.FIO
    Call selectPositionOption(curItem.PositionStatus)
End Sub


Private Sub ButtonClose_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "�������" - ������� � ����� ������ ��
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Unload Me
    UKForm.Show
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ���������� ����������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New employee_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub ButtonChange_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ��������� ����������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ������� ������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub ButtonDel_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ �������� ���������
' Last update: 28.03.2018
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
        curMC.initialEmployees
    End If
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' ����������/��������� ����������
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    Dim ps As PositionStatusEnum
    
    If formNotFill Then
        MsgBox "��������� �� ��� ����!", vbInformation + vbOKOnly, "������"
                                                            
        GoTo cleanHandler
    End If
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        If Me.SelectPositionChiefEngineer Then
            ps = PositionStatusEnum.psChiefEngineer
        ElseIf Me.SelectPositionDirector Then
            ps = PositionStatusEnum.psDirector
        ElseIf Me.SelectPositionOther Then
            ps = PositionStatusEnum.psOther
        End If
        On Error GoTo errHandler
        curItem.update newFirstName:=Me.TBFirstName.Value, _
                        newSecondName:=Me.TBSecondName.Value, _
                        newLastName:=Me.TBLastName.Value, _
                        newMcId:=curMC.Id, _
                        newPosition:=Me.TBPosition.Value, _
                        newSign:=Me.TBSign.Value, _
                        newPositionStatus:=ps, _
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
' �������� �� ���������� ���� ����������� �����
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    If Trim(Me.TBFirstName.Value) <> "" And Trim(Me.TBLastName.Value) <> "" _
                                    And Trim(Me.TBPosition.Value) <> "" And _
                                    Trim(Me.TBSecondName.Value) <> "" Then
        formNotFill = False
    Else
        formNotFill = True
    End If
End Function


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' ���������� ������ �����������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As ListItem
    
    ' ���������� �������
    Me.ListViewList.ListItems.Clear
    For i = 1 To curMC.employees.count
        Set curItem = curMC.employees(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormEmployeeEnum.feMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormEmployeeEnum.feFirstName).text = _
                                                            curItem.FirstName
        listX.ListSubItems(FormEmployeeEnum.feLastName).text = _
                                                            curItem.LastName
        listX.ListSubItems(FormEmployeeEnum.fePosition).text = _
                                                            curItem.Position
        listX.ListSubItems(FormEmployeeEnum.feSecondName).text = _
                                                            curItem.SecondName
        listX.ListSubItems(FormEmployeeEnum.feSignReport).text = _
                                            BoolToYesNo(curItem.ReportSign)
        
    Next i
    
    ' �������������� ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set listX = Nothing
    Set curItem = Nothing
End Sub


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' ������� ��������� �����
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Me.TBFirstName.Value = ""
    Me.TBLastName.Value = ""
    Me.TBPosition.Value = ""
    Me.TBSecondName.Value = ""
    Me.TBSign.Value = False
    Me.SelectPositionOther = True
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub


Private Sub selectPositionOption(Value As PositionStatusEnum)
' ----------------------------------------------------------------------------
' ��������� ������ ���������
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    If Value = PositionStatusEnum.psDirector Then
        Me.SelectPositionDirector = True
    ElseIf Value = PositionStatusEnum.psChiefEngineer Then
            Me.SelectPositionChiefEngineer = True
    ElseIf Value = PositionStatusEnum.psOther Then
        Me.SelectPositionOther = True
    End If
    posId = Value
   
End Sub
