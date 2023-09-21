VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlanWorkForm 
   Caption         =   "���� ����������� ������"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   OleObjectBlob   =   "PlanWorkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlanWorkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------
' ���������� �����
' ----------------------------------------------------------------------------
Public bldnId As Long               ' ��� ����
Public workId As Long               ' ��� ���������� ������
Private changedWork As plan_work_class ' ���������� ������ (���� ���������)
Public mcId As Long

Private WithEvents m_Cal As cCalendar
Attribute m_Cal.VB_VarHelpID = -1


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� �����, ���������� ���������� �������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    ' ���������� �������
    Call reloadComboBox(rcmContractor, Me.ComboBoxContractor)
    Call reloadComboBox(rcmGWT, Me.ComboBoxGlobalWorkType)
    Call reloadComboBox(rcmWorkType, Me.ComboBoxWorkType)
    Call reloadComboBox(rcmPlanStatuses, Me.ComboBoxStatus)
    Call reloadMonths
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����, ���� �������� ������, �� ��������� ����
' 25.08.2021
' ----------------------------------------------------------------------------
    If workId > 0 Then
        Set changedWork = New plan_work_class
        changedWork.initial workId
        Call selectComboBoxValue(Me.ComboBoxMonth, Month(changedWork.WorkDate))
        Call selectComboBoxValue(Me.ComboBoxYear, Year(changedWork.WorkDate))
        Call selectComboBoxValue(Me.ComboBoxGlobalWorkType, _
                                            changedWork.GWT.Id)
        Call selectComboBoxValue(Me.ComboBoxWorkType, _
                                            changedWork.WorkKind.workType.Id)
        Call selectComboBoxValue(Me.ComboBoxWorkKind, _
                                            changedWork.WorkKind.Id)
        Call selectComboBoxValue(Me.ComboBoxContractor, _
                                            changedWork.Contractor.Id)
        Call selectComboBoxValue(Me.ComboBoxStatus, _
                                            changedWork.Status.Id)
        Me.TextBoxNote = changedWork.Note
        Me.TextBoxPrivateNote = changedWork.PrivateNote
        Me.TextBoxSum = CDbl(changedWork.sum)
        Me.TextBoxSmetaSum = dblValue(changedWork.smetaSum)
        Me.TextBoxEmployee = changedWork.Employee
        Me.LabelPickDateBegin = changedWork.beginDate
        Me.LabelPickDateEnd = changedWork.EndDate
    Else
        ' ��� �������� ������ ����� �������� ������ ������������ �������
        Call reloadComboBox(rcmPlanStatusesNewWork, Me.ComboBoxStatus, _
                                            defValue:=plan_statuses.PWPlan)
        Me.LabelPickDateBegin = NOTDATE
        Me.LabelPickDateEnd = NOTDATE
    End If
        
End Sub


Private Sub ComboBoxStatus_Change()
' ----------------------------------------------------------------------------
' ��������, ����� �� ������������ ���� ������ ����
' 25.08.2021
' ----------------------------------------------------------------------------
    If Me.ComboBoxStatus > -1 Then
        Me.LabelPickDateBegin.Enabled = (Me.ComboBoxStatus.Value = plan_statuses.PWInWork)
        Me.LabelPickDateEnd.Enabled = (Me.ComboBoxStatus.Value = plan_statuses.PWInWork)
        Me.BtnPickDateBegin.Enabled = (Me.ComboBoxStatus.Value = plan_statuses.PWInWork)
        Me.BtnPickDateEnd.Enabled = (Me.ComboBoxStatus.Value = plan_statuses.PWInWork)
    End If
End Sub


Private Sub ComboBoxWorkType_Change()
' ----------------------------------------------------------------------------
' ��� ��������� ���� ����� ����������� ���� ������.
' Last update: 22.04.2016
' ----------------------------------------------------------------------------
    If Me.ComboBoxWorkType.ListIndex > -1 Then
        Call reloadComboBox(rcmWorkKind, Me.ComboBoxWorkKind, _
                                        initValue:=Me.ComboBoxWorkType.Value)
    End If
End Sub


Private Sub BtnPickDateBegin_Click()
' ----------------------------------------------------------------------------
' ����� ���� ������ ������
' 26.08.2021
' ----------------------------------------------------------------------------
    PickDateToLabel Me.LabelPickDateBegin, _
            changedWork, _
            "���� ������ ������"
End Sub


Private Sub BtnPickDateEnd_Click()
' ----------------------------------------------------------------------------
' ����� ���� ��������� ������
' 26.08.2021
' ----------------------------------------------------------------------------
    PickDateToLabel Me.LabelPickDateEnd, _
            changedWork, _
            "���� ��������� ������"
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' ��������� ������� ������ "���������"
' 25.08.2021
' ----------------------------------------------------------------------------
    Dim curPlanWork As New plan_work_class
    Dim wId As Long
    
    If Not isPlanFill Then
        Call setMsg("��������� �� ��� ����������� ����")
        Exit Sub
    End If
    
    On Error GoTo errHandler
    ' ���������� ������
    If Not changedWork Is Nothing Then
        ' ������ � ������ ��������� ������ "���������",
        ' ������ ����� ������ �������� ����� ������
        If changedWork.Status.Id <> plan_statuses.PWDone And _
                        Me.ComboBoxStatus.Value = plan_statuses.PWDone Then
            MsgBox "����� �������� ���������� ������" & vbCrLf & _
                            "������� �� ��������������� ������", _
                            vbExclamation, "������ ����������"
            Exit Sub
        End If
        wId = NOTVALUE
        changedWork.update newGWT:=Me.ComboBoxGlobalWorkType.Value, _
                        newWKind:=Me.ComboBoxWorkKind.Value, _
                        newDate:=DateSerial(Me.ComboBoxYear.Value, _
                                                Me.ComboBoxMonth.Value, 1), _
                        newSum:=dblValue(Me.TextBoxSum.Value), _
                        newSmetaSum:=dblValue(Me.TextBoxSmetaSum.Value), _
                        newNote:=Me.TextBoxNote.Value, _
                        newPrivateNote:=Trim(Me.TextBoxPrivateNote), _
                        newContractor:=Me.ComboBoxContractor.Value, _
                        newStatus:=Me.ComboBoxStatus.Value, _
                        newEmployee:=Me.TextBoxEmployee.Value, _
                        newBeginDate:=CDate(Me.LabelPickDateBegin), _
                        newEndDate:=CDate(Me.LabelPickDateEnd)
    Else

        Set changedWork = New plan_work_class
        changedWork.create bldnId:=bldnId, _
                        gwtId:=Me.ComboBoxGlobalWorkType.Value, _
                        workKindID:=Me.ComboBoxWorkKind.Value, _
                        WorkDate:=DateSerial(Me.ComboBoxYear.Value, _
                                            Me.ComboBoxMonth.Value, 1), _
                        workSum:=dblValue(Me.TextBoxSum.Value), _
                        smetaSum:=dblValue(Me.TextBoxSmetaSum.Value), _
                        workNote:=Trim(Me.TextBoxNote.Value), _
                        workPrivateNote:=Trim(Me.TextBoxPrivateNote.Value), _
                        contractorId:=Me.ComboBoxContractor.Value, _
                        mcId:=mcId, _
                        Status:=Me.ComboBoxStatus.Value, _
                        Employee:=Me.TextBoxEmployee.Value
        
    End If
    
    BuildingForm.planWorkChanged = True
    Unload Me
    BuildingForm.show
    GoTo cleanHandler

errHandler:
    Call setMsg(Err.Description, True)
cleanHandler:
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' About: ��������� ������� ������ "������"
' Last update: 07.06.2016
' ----------------------------------------------------------------------------
    Unload Me           ' ����������� �����
    BuildingForm.workChanged = False
    BuildingForm.show   ' ����� ����� ������
End Sub


Private Sub reloadMonths()
' ----------------------------------------------------------------------------
' ���������� ������� ����������� �����
' Last update: 20.02.2018
' ----------------------------------------------------------------------------
    Dim curListIndex As Long, i As Long
    With Me.ComboBoxMonth
        .ColumnCount = 2
        .BoundColumn = ComboColumns.ccId + 1
        .Clear
        curListIndex = 0
        For i = 1 To 12
            .AddItem
            .list(curListIndex, ComboColumns.ccname) = MonthName(i)
            .list(curListIndex, ComboColumns.ccId) = i
            curListIndex = curListIndex + 1
        Next i
        .ColumnWidths = ";0"
    End With
    With Me.ComboBoxYear
        .ColumnCount = 2
        .BoundColumn = ComboColumns.ccId + 1
        .Clear
        curListIndex = 0
        For i = FIRST_PLAN_YEAR To Year(Now) + 3
            .AddItem
            .list(curListIndex, ComboColumns.ccname) = i
            .list(curListIndex, ComboColumns.ccId) = i
            curListIndex = curListIndex + 1
        Next i
        .ColumnWidths = ";0"
    End With
End Sub


Private Sub CommandButtonDone_Click()
' ----------------------------------------------------------------------------
' ������� ���������� ������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Set WorkForm.planWork = changedWork
    Unload Me
    WorkForm.show
End Sub


Private Sub moveCalendarToControl(curControl As MSForms.Control)
    If m_Cal Is Nothing Then Exit Sub
    
    Dim vLeft As Long, vTop As Long
    
    vLeft = curControl.Left + curControl.Width ' + Me.Left
    vTop = curControl.Top + curControl.Height ' + Me.Top
    m_Cal.Move vLeft, vTop
End Sub


Private Function isPlanFill() As Boolean
' ----------------------------------------------------------------------------
' ��������� �� ��� ����������� ��� ����� ����
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    isPlanFill = False
    If Me.ComboBoxContractor.ListIndex = -1 Then Exit Function
    If Me.ComboBoxMonth.ListIndex = -1 Then Exit Function
    If Me.ComboBoxGlobalWorkType.ListIndex = -1 Then Exit Function
    If Me.ComboBoxWorkKind.ListIndex = -1 Then Exit Function
    If Me.ComboBoxWorkType.ListIndex = -1 Then Exit Function
    If Me.ComboBoxYear.ListIndex = -1 Then Exit Function
    isPlanFill = True
End Function


Private Sub setMsg(msgText As String, Optional isError As Boolean = True)
' ----------------------------------------------------------------------------
' ����� ���������
' Last update: 12.11.2018
' ----------------------------------------------------------------------------
    Me.LabelMsg.Caption = msgText
    Me.LabelMsg.ForeColor = IIf(isError, RGB(255, 0, 0), RGB(0, 0, 0))
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' ������ �������� ����� ���������, �.�. ����� �����
'        ����������� �������� ����� ����� ���
' Last update: 01.03.2016
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


Private Sub UserForm_Deactivate()
' ----------------------------------------------------------------------------
' ����������� ���������� ��� ��������
' Last update: 20.02.2018
' ----------------------------------------------------------------------------
    Set changedWork = Nothing
End Sub
