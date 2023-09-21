VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkMaintenanceForm 
   Caption         =   "������ �� ����������"
   ClientHeight    =   10665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   OleObjectBlob   =   "WorkMaintenanceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkMaintenanceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curWork As work_maintenance
Public curBldn As building_class
Public isRecalc As Boolean          ' ����, ����� �� ������������� ������ ���
Private isFirst As Boolean          ' ���� ������ ��������
Private curMaterial As Long         ' ��� �������� ���������


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ���������� ��������� ��� ������������� �����
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    Dim i As Long
    
    isRecalc = False
    Call reloadComboBox(rcmTermDESC, Me.ComboBoxTerms)
    Call reloadComboBox(rcmWorkType, Me.ComboBoxWorkType)
    
    With Me.ListViewMaterials
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        .Gridlines = True       ' ����������� ������
        With .ColumnHeaders
            For i = 1 To FormWorkMaterialsEnum.fwmMax
                .add
            Next i
            .Item(FormWorkMaterialsEnum.fwmMaterialCost + 1).text = "����"
            .Item(FormWorkMaterialsEnum.fwmMaterialCount + 1).text = "���-��"
            .Item(FormWorkMaterialsEnum.fwmMaterialId + 1).text = "���"
            .Item(FormWorkMaterialsEnum.fwmMaterialName + 1).text = "��������"
            .Item(FormWorkMaterialsEnum.fwmMaterialNote + 1).text = "���������"
            .Item(FormWorkMaterialsEnum.fwmMaterialSi + 1).text = "��.���."
            .Item(FormWorkMaterialsEnum.fwmMaterialSum + 1).text = "�����"
            
            .Item(FormWorkMaterialsEnum.fwmMaterialId + 1).Width = 0
            .Item(FormWorkMaterialsEnum.fwmMaterialCost + 1).Width = 40
            .Item(FormWorkMaterialsEnum.fwmMaterialName + 1).Width = 150
            .Item(FormWorkMaterialsEnum.fwmMaterialCount + 1).Width = 40
            .Item(FormWorkMaterialsEnum.fwmMaterialNote + 1).Width = 50
            .Item(FormWorkMaterialsEnum.fwmMaterialSi + 1).Width = 50
        End With
    End With
    
    isFirst = True
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����, ���������� �����, ���� ���������
' 13.02.2023
' ----------------------------------------------------------------------------
    On Error GoTo errHandler

    If isFirst Then
        ' ���� ��� �� ������, �� ��������� �����
        If curBldn Is Nothing Then
            MsgBox "�� ������ ���", vbCritical, "������"
            Unload Me
            Exit Sub
        End If
        If curBldn.Id = NOTVALUE Then
            MsgBox "�� ������ ���", vbCritical, "������"
            Unload Me
            Exit Sub
        End If
        
        Me.LabelInfo.Caption = curBldn.Address & vbCr & _
                IIf(curWork Is Nothing, "���������� ������", "��������� ������")
        
        ' ���������� �����
        If Not curWork Is Nothing Then
            Call selectComboBoxValue(Me.ComboBoxWorkType, _
                                                    curWork.WorkKind.workType.Id)
            Call selectComboBoxValue(Me.ComboBoxWorkKind, curWork.WorkKind.Id)
            Me.TextBoxManHour.Value = curWork.ManHours
            Me.TextBoxNote.Value = curWork.Note
            Me.TextBoxPrivateNote.Value = curWork.PrivateNote
            Me.CheckBoxUsedWork = Not curWork.PrintFlag
        Else
            Set curWork = New work_maintenance
        End If
        
        Call fillDateBox    ' ����� ����
        Call fillManHourCost    ' ��������� ������������
    End If
    If Not curWork Is Nothing Then Call fillMaterials
    Call recalcWorkSum
    isFirst = False
        
errHandler:
    If Err.Number <> 0 Then Me.LabelError.Caption = Err.Description
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' ������ �������� ����� ���������,
'                    �.�. ����� ����� ����������� �������� ����� ����� ���
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' �������� �����, ����� ����������
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    Set curWork = Nothing
    Set curBldn = Nothing
End Sub


Private Sub ListViewMaterials_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��� ������ ��������� ��������� ��� ������
' Last update: 17.10.2019
' ----------------------------------------------------------------------------
    curMaterial = Item.index
End Sub


Private Sub ComboBoxTerms_Change()
' ----------------------------------------------------------------------------
' ��� ����� ���� ��������� ���� ��������� ������������
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    If isRecalc Then
        Call fillManHourCost
    End If
End Sub


Private Sub ComboBoxWorkType_Change()
' ----------------------------------------------------------------------------
' ��� ������ ���� ������ ��������� ������ �����
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    If Me.ComboBoxWorkType.ListIndex > -1 Then
        Call reloadComboBox(rcmWorkKind, Me.ComboBoxWorkKind, _
                                                    Me.ComboBoxWorkType.Value)
    End If
End Sub


Private Sub TextBoxTransport_Change()
' ----------------------------------------------------------------------------
' ��� ��������� ����� ������������� ��������� ������
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    Call recalcWorkSum
End Sub


Private Sub TextBoxManHour_Change()
' ----------------------------------------------------------------------------
' ��� ��������� ����� ������������� ��������� ������
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    Call recalcWorkSum
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' ���������� ���������
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    With WorkMaterialForm
        Set .curWork = curWork
        .Caption = "���������� ���������"
        .Show
    End With
    Call fillMaterials
End Sub


Private Sub ButtonChange_Click()
' ----------------------------------------------------------------------------
' ��������� ���������
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    With WorkMaterialForm
        Set .curWork = curWork
        .materialIdx = curMaterial
        .Caption = "��������� ���������"
        .Show
    End With
    Call fillMaterials
End Sub


Private Sub ButtonDelete_Click()
' ----------------------------------------------------------------------------
' �������� ���������
' Last update: 17.10.2019
' ----------------------------------------------------------------------------
    If curMaterial <> 0 Then
        With curWork.Materials(curMaterial)
            If ConfirmDeletion(material_types(CStr(.MaterialId)).Name & _
                                " " & .MaterialNote & _
                                " " & .MaterialCost & _
                                " " & .MaterialCount & _
                                " " & .MaterialSi) Then
                curWork.Materials.remove curMaterial
                Call fillMaterials
            End If
        End With
    End If
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' ���������� ������
' 13.02.2023
' ----------------------------------------------------------------------------

    On Error GoTo errHandler

    If formCheck Then
        If getManHourCost = NOTVALUE Then
            MsgBox "��� ������� ���������� �� ����������� ��������� " & _
                    "������������ � ��������� ������", vbCritical, "������"
            Exit Sub
        End If
        If curWork.Id = NOTVALUE Then
            curWork.create newDate:=Me.ComboBoxTerms.Value, _
                            newWorkKind:=Me.ComboBoxWorkKind.Value, _
                            newManHours:=dblValue(Me.TextBoxManHour), _
                            newNote:=Me.TextBoxNote, _
                            newPrintFlag:=Not Me.CheckBoxUsedWork, _
                            newBldnId:=curBldn.Id, _
                            newPrivateNote:=Me.TextBoxPrivateNote
            Call saveUseWorkDate(Me.ComboBoxTerms.Value)
        Else
            curWork.update newDate:=Me.ComboBoxTerms.Value, _
                            newWorkKind:=Me.ComboBoxWorkKind.Value, _
                            newManHours:=dblValue(Me.TextBoxManHour), _
                            newNote:=Me.TextBoxNote, _
                            newPrintFlag:=Not Me.CheckBoxUsedWork, _
                            newBldnId:=curBldn.Id, _
                            newPrivateNote:=Me.TextBoxPrivateNote
        End If
        Call formClose(True)
    Else
        MsgBox "��������� �� ��� ����������� ����", vbExclamation, "��������"
    End If
    GoTo cleanHandler

errHandler:
    MsgBox Err.Description, vbCritical, "������"
    
cleanHandler:
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' ����� ��� ����������
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    Call formClose(False)
End Sub


Private Sub recalcWorkSum()
' ----------------------------------------------------------------------------
' �������� ��������� ������
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    Dim tmpSum As Currency
    
    tmpSum = dblValue(Me.TextBoxManHour) * dblValue(Me.LabelManHourCost) + _
                dblValue(curWork.MaterialCost)
    Me.LabelWorkSum = "��������� ��������� ������: " & tmpSum & " ���."
End Sub


Private Function formCheck() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������� �����
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    formCheck = (Me.ComboBoxWorkKind.ListIndex > -1) And _
                (Me.ComboBoxWorkType.ListIndex > -1)
End Function


Private Sub fillDateBox()
' ----------------------------------------------------------------------------
' ���������� ���� � �����
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    Dim tmpDateTerm As Long
    
    isRecalc = True
    If curWork.Id = NOTVALUE Then
        tmpDateTerm = getUseWorkDate()
        If tmpDateTerm <> NOTVALUE Then
            Call selectComboBoxValue(Me.ComboBoxTerms, tmpDateTerm)
        Else
            Me.ComboBoxTerms.ListIndex = 0
        End If
    Else
        Call selectComboBoxValue(Me.ComboBoxTerms, curWork.WorkDate.Id)
    End If
End Sub


Private Sub fillManHourCost()
' ----------------------------------------------------------------------------
' ���������� ��������� ������������
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    Me.LabelManHourCost.Caption = Application.Max(getManHourCost, 0)
End Sub


Private Function getManHourCost() As Currency
' ----------------------------------------------------------------------------
' ��������� ������������
' 23.03.2022
' ----------------------------------------------------------------------------
    Dim curManHourCost As New man_hour_cost
    
    getManHourCost = NOTVALUE
    
    If Not curWork.Id = NOTVALUE Then
        curManHourCost.initial curWork.Contractor.Id, _
                                Me.ComboBoxTerms.Value, _
                                curWork.ManHourCostMode.Id
    Else
        curManHourCost.initialByBldnTerm curBldn.Id, _
                                Me.ComboBoxTerms.Value
    End If
        
    On Error Resume Next
    getManHourCost = curManHourCost.Cost
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
    Set curManHourCost = Nothing
End Function


Private Sub fillMaterials()
' ----------------------------------------------------------------------------
' ���������� ����������
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    If Not curWork Is Nothing Then
        Dim i As Long, j As Long
        Dim listX As ListItem
        Dim curItem As works_materials
            
        curMaterial = 0
        Me.ListViewMaterials.ListItems.Clear
        
        For i = 1 To curWork.Materials.count
            Set curItem = curWork.Materials(i)
            Set listX = Me.ListViewMaterials.ListItems.add(, , curItem.Id)
            With listX.ListSubItems
                For j = 1 To FormWorkMaterialsEnum.fwmMax - 1
                    .add
                Next j
                .Item(FormWorkMaterialsEnum.fwmMaterialCost).text = _
                                                        curItem.MaterialCost
                .Item(FormWorkMaterialsEnum.fwmMaterialCount).text = _
                                                        curItem.MaterialCount
                .Item(FormWorkMaterialsEnum.fwmMaterialName).text = _
                            material_types(CStr(curItem.MaterialId)).Name
                .Item(FormWorkMaterialsEnum.fwmMaterialNote).text = _
                                                        curItem.MaterialNote
                .Item(FormWorkMaterialsEnum.fwmMaterialSi).text = _
                                                        curItem.MaterialSi
                .Item(FormWorkMaterialsEnum.fwmMaterialSum).text = _
                                                        curItem.MaterialSum
            End With
        Next i
        Call recalcWorkSum
    End If
End Sub


Private Sub formClose(Changed As Boolean)
' ----------------------------------------------------------------------------
' �������� �����
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    Unload Me
    BuildingForm.workChanged = Changed
    BuildingForm.Show
End Sub
