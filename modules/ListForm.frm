VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListForm 
   Caption         =   "�������� ���������"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14910
   OleObjectBlob   =   "ListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fType As ListFormType       ' ��� ����������
Private curItem As MSComctlLib.listItem
Private FormKeys As New Dictionary  ' ��������� ���������


Property Get formType() As ListFormType
' ----------------------------------------------------------------------------
' ������� ���� ����������
' Last update: 07.05.2019
' ----------------------------------------------------------------------------
    formType = fType
End Property


Property Let formType(newType As ListFormType)
' ----------------------------------------------------------------------------
' ��������� ���� ����������
' Last update: 07.05.2019
' ----------------------------------------------------------------------------
    fType = newType
End Property


Public Function setParameter(paramKey As String, paramValue As Variant)
' ----------------------------------------------------------------------------
' ��������� ��������� ��� �����
' 02.08.2021
' ----------------------------------------------------------------------------
    If FormKeys.Exists(paramKey) Then
        FormKeys(paramKey) = paramValue
    Else
        FormKeys.add paramKey, paramValue
    End If
End Function


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� ����� - ����� ���������
' 23.03.2022
' ----------------------------------------------------------------------------
    With Me.ListViewList
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
    End With
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����
' 19.10.2022
' ----------------------------------------------------------------------------

    ' ���� ��� ���������� �� �����, �� �����
    If formType = 0 Then Unload Me
    
    ' ��� ������ ���� ������ ���������� - ������� � ����
    If formType = lftTmpCounters Then
        Me.BtnAdd.Caption = "�������"
        Me.BtnChange.Visible = False
        Me.BtnDelete.Visible = False
    End If
    
    ' ���� ������� �������� �������� ������, �� ��� ��������
    If FormKeys.Exists("HideButtons") Then
        If FormKeys("HideButtons") Then
            Me.BtnAdd.Visible = False
            Me.BtnChange.Visible = False
            Me.BtnDelete.Visible = False
        End If      ' hidebutton
    End If          ' key hidebutton exists
    
    ' ������ ��������� � ������� ��������� ������������
    If formType = lftManHourCost Then
        Me.BtnAdd.Visible = False
        Me.BtnDelete.Visible = False
    End If
    
    ' �������� ����������� ��������� ������
    If FormKeys.Exists("HideButtonAdd") Then
        If FormKeys("HideButtonAdd") Then Me.BtnAdd.Visible = False
    End If
    If FormKeys.Exists("HideButtonChange") Then
        If FormKeys("HideButtonChange") Then Me.BtnChange.Visible = False
    End If
    If FormKeys.Exists("HideButtonDelete") Then
        If FormKeys("HideButtonDelete") Then Me.BtnDelete.Visible = False
    End If
    
    Call reloadList
End Sub


Private Sub ListViewList_ColumnClick( _
                            ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'-----------------------------------------------------------------------------
' ���������� ������ ��� ������ �� ���������
' 25.09.2019
'-----------------------------------------------------------------------------
    Me.ListViewList.Sorted = True
    Me.ListViewList.SortKey = ColumnHeader.index - 1
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.listItem)
' ----------------------------------------------------------------------------
' ���������� ���������� ��������
' 23.03.2022
' ----------------------------------------------------------------------------
    Set curItem = Item
End Sub


Private Sub BtnAdd_Click()
'-----------------------------------------------------------------------------
' ���������� ������ ��������
' 19.10.2022
'-----------------------------------------------------------------------------
    If Me.formType = lftChairmansSign Then
        Call CreateItem
    Else
        Call UpdateItem(isChange:=False)
    End If
End Sub


Private Sub BtnChange_Click()
'-----------------------------------------------------------------------------
' ��������� ��������
' 27.09.2021
'-----------------------------------------------------------------------------
    If curItem Is Nothing Then Exit Sub
    
    Call UpdateItem(isChange:=True)
End Sub


Private Sub BtnDelete_Click()
'-----------------------------------------------------------------------------
' �������� ��������
' 19.10.2022
'-----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    If curItem Is Nothing Then Exit Sub
    
    Dim curObject As Object
    If Me.formType = lftExpenseGroups Or Me.formType = lftExpenseItems Or _
            Me.formType = lftCommonPropertyGroup Or _
            Me.formType = lftCommonPropertyElement Or _
            Me.formType = lftCommonPropertyParameter Or _
            Me.formType = lftManHourCostModes _
        Then
'        Dim curObject As base_form_class
        Set curObject = SetCurObject1()
    ElseIf Me.formType = lftChairmansSign Then
        ' ��� ��������� � ������� �������� ������ ��� ��������
        ' ����� ���� ���� ���������
        Dim tmpCurObject As New base_delete_class
        Set tmpCurObject = New bldn_chairman_sign
        tmpCurObject.initial curItem
        Set curObject = tmpCurObject
        Set tmpCurObject = Nothing
    Else
'        Dim curObject As basicIdNameClass
        Set curObject = SetCurObject()
    End If
    If Me.formType <> lftChairmansSign Then
        curObject.initial CLng(curItem)
    End If
    
    If ConfirmDeletion(curObject.Name) Then
        curObject.delete
    End If
    Set curObject = Nothing
    Call reloadList

errHandler:
    If errorHasChildren(Err.Description) Then
        MsgBox "���� ����������� �������, �������� ���������", _
                                            vbExclamation, "������ ��������"
    ElseIf errorHasNoPrivilegies(Err.Description) Then
        MsgBox "�� ������� ����", vbExclamation, "������ ��������"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & vbCr & Err.Source & vbCr & Err.Description, _
                                                vbCritical, "������ ��������"
    End If
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' �������� �����
' Last update: 17.05.2019
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonPrint_Click()
' ----------------------------------------------------------------------------
' ������� �����������
' 08.10.2021
' ----------------------------------------------------------------------------
    Call ExportCurrentList
End Sub


Private Sub reloadList()
'-----------------------------------------------------------------------------
' ���������� ������, ����� ������� ������
' 03.06.2022
'-----------------------------------------------------------------------------
    Me.ListViewList.Visible = False
    Select Case formType
        Case lftCounterModels:
            Me.Caption = "������ ������� ���������"
            Call reloadCounterModels
        Case lftWorkMaterialTypes
            Me.Caption = "���������"
            Call reloadWorkMaterialTypes
        Case lftFinanceSources
            Me.Caption = "��������� ��������������"
            Call reloadFinanceSources
        Case lftMunicipalDistrict
            Me.Caption = "M������������ �����������"
            Call reloadMunicipalDistrict
        Case ListFormType.lftHCounterPartTypes
            Me.Caption = "����� ���� ��������������"
            Call reloadHCounterPartTypes
        Case ListFormType.lftTmpCounters
            Me.Caption = "���� ������� �������� �����"
            Call reloadTmpCounters
            Me.BtnAdd.SetFocus
        Case Else
            Call fillList
    End Select
    Me.Caption = Me.Caption & ". ������ " & AppConfig.DBServer
    Me.ListViewList.Visible = True
    Me.ListViewList.Visible = False
    Me.ListViewList.Visible = True
End Sub


Private Sub fillList()
'-----------------------------------------------------------------------------
' ���������� ������, ����� ������� ������
' 19.10.2022
'-----------------------------------------------------------------------------
    Dim curList As basicListForm
    Dim tmpList As Object
    
    Select Case formType
        Case ListFormType.lftExpenseItems
            Set curList = expense_items
        Case ListFormType.lftBldnSubaccounts
            Set tmpList = New subaccounts
            If FormKeys.Exists("bldnId") Then
                tmpList.initialByBldn (FormKeys("bldnId"))
            End If
            Set curList = tmpList
        Case ListFormType.lftFlatHistory
            Set tmpList = New flats
            If FormKeys.Exists("flatId") Then
                tmpList.initialFlatHistory longValue(FormKeys("flatId"))
            End If
            Set curList = tmpList
        Case ListFormType.lftContractors
            Set curList = New contractor_list
        Case ListFormType.lftCommonPropertyGroup
            Set curList = common_property_groups
        Case ListFormType.lftExpenseGroups
            Set curList = New expense_groups
        Case ListFormType.lftCommonPropertyElement
            Set curList = New common_property_elements
        Case ListFormType.lftCommonPropertyParameter
            Set curList = New common_property_parameters
        Case ListFormType.lftManHourCostModes
            Set tmpList = man_hour_cost_modes
            tmpList.reload
            Set curList = tmpList
        Case ListFormType.lftManHourCost
            Set tmpList = New man_hour_costs
            tmpList.initialCurrent
            Set curList = tmpList
        Case ListFormType.lftFlatAccrueds
            Set tmpList = New flat_accrueds
            tmpList.initial FormKeys("flatNo"), longValue(FormKeys("flatId"))
            Set curList = tmpList
        Case ListFormType.lftChairmansSign
            Set tmpList = New bldn_chairmans_sign
            If FormKeys.Exists("bldnId") Then
                tmpList.initial (FormKeys("bldnId"))
            End If
            Set curList = tmpList
    End Select
    Set tmpList = Nothing
    
    Call curList.fillListform(Me.ListViewList)
    Me.Caption = curList.Title
    Set curList = Nothing
End Sub


Private Sub CreateItem()
'-----------------------------------------------------------------------------
' ���������� �������
' 19.10.2022
'-----------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Dim curObject As base_create_class
    Dim initParams As New Dictionary
    If Me.formType = lftChairmansSign Then
        initParams.add "BldnId", FormKeys("bldnId")
        Set curObject = New bldn_chairman_sign
        curObject.initial initParams
        
    End If
    
    curObject.showForm
    
    Set curObject = Nothing
    
    Call reloadList
    
errHandler:
    If errorHasChildren(Err.Description) Then
        MsgBox "���� ����������� �������, �������� ���������", _
                                            vbExclamation, "������ ��������"
    ElseIf errorHasNoPrivilegies(Err.Description) Then
        MsgBox "�� ������� ����", vbExclamation, "������ ��������"
    ElseIf errorHasNoValues(Err.Number) Then
        MsgBox "������ �� �����", vbExclamation, "������ ��������"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & vbCr & Err.Source & vbCr & Err.Description, _
                                                vbCritical, "������ ��������"
    End If

End Sub

Private Sub UpdateItem(isChange As Boolean)
'-----------------------------------------------------------------------------
' ���������/���������� �������
' 19.10.2022
'-----------------------------------------------------------------------------

    On Error GoTo errHandler
    
    Dim curObject As Object
    If Me.formType = lftExpenseGroups Or Me.formType = lftExpenseItems Or _
            Me.formType = lftCommonPropertyGroup Or _
            Me.formType = lftCommonPropertyElement Or _
            Me.formType = lftCommonPropertyParameter Or _
            Me.formType = lftManHourCostModes Or _
            Me.formType = lftManHourCost Or _
            Me.formType = lftChairmansSign _
                                        Then
'        Dim curObject As base_form_class
        Set curObject = SetCurObject1()
    Else
'        Dim curObject As basicIdNameClass
        Set curObject = SetCurObject()
    End If
    
    If isChange Then
        If Not curItem Is Nothing Then
            curObject.initial CLng(curItem)
        End If
    End If
    
    curObject.showForm isChange:=isChange
    
    Set curObject = Nothing
    Set curItem = Nothing
    
    Call reloadList
    
errHandler:
    If errorHasChildren(Err.Description) Then
        MsgBox "���� ����������� �������, �������� ���������", _
                                            vbExclamation, "������ ��������"
    ElseIf errorHasNoPrivilegies(Err.Description) Then
        MsgBox "�� ������� ����", vbExclamation, "������ ��������"
    ElseIf errorHasNoValues(Err.Number) Then
        MsgBox "������ �� �����", vbExclamation, "������ ��������"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & vbCr & Err.Source & vbCr & Err.Description, _
                                                vbCritical, "������ ��������"
    End If

End Sub


Private Sub reloadCounterModels()
'-----------------------------------------------------------------------------
' ���������� ������ ������� �������� �����
' Last update: 29.05.2019
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curLItem As counter_model
    Dim CIWIDTH As Integer, DTIWIDTH As Integer, NAMEWIDTH As Integer
    
    
    With Me.ListViewList
        CIWIDTH = 135
        DTIWIDTH = 70
        NAMEWIDTH = .Width - CIWIDTH - DTIWIDTH - 5
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormCounterModels.cmfMax
                .add
            Next i
            .Item(FormCounterModels.cmfId + 1).text = "���"
            .Item(FormCounterModels.cmfHasDTI + 1).text = "������� ���"
            .Item(FormCounterModels.cmfCI + 1).text = "������������� ��������"
            .Item(FormCounterModels.cmfName + 1).text = "��������"
        End With
        .ColumnHeaders(FormCounterModels.cmfId + 1).Width = 0
        .ColumnHeaders(FormCounterModels.cmfCI + 1).Width = CIWIDTH
        .ColumnHeaders(FormCounterModels.cmfHasDTI + 1).Width = DTIWIDTH
        .ColumnHeaders(FormCounterModels.cmfName + 1).Width = NAMEWIDTH
        
        ' ���������� �������
        .ListItems.Clear
        counter_models.reload
        For i = 1 To counter_models.count
            Set curLItem = counter_models(i)
            Set listX = .ListItems.add(, , curLItem.Id)
            For j = 1 To FormCounterModels.cmfMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormCounterModels.cmfName).text = curLItem.Name
            listX.ListSubItems(FormCounterModels.cmfCI).text = _
                                                curLItem.CalibrationInterval
            listX.ListSubItems(FormCounterModels.cmfHasDTI).text = _
                                                BoolToYesNo(curLItem.HasDTI)
        Next i
    End With
    
    ' ������ ��������
'    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set curLItem = Nothing
    Set listX = Nothing
End Sub


Private Sub reloadWorkMaterialTypes()
'-----------------------------------------------------------------------------
' ���������� ����������
' Last update: 29.10.2019
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curList As material_types
    Dim curItem As material_type
    
    On Error GoTo errHandler
    
    With Me.ListViewList
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormWorkMaterialType.fwmtMax
                .add
            Next i
            .Item(FormWorkMaterialType.fwmtMaterialId + 1).text = "���"
            .Item(FormWorkMaterialType.fwmtMaterialName + 1).text = "��������"
            .Item(FormWorkMaterialType.fwmtIsTransport + 1).text = "���������"
        End With
        
        ' ���������� �������
        .ListItems.Clear
        Set curList = material_types
        curList.reload
        For i = 1 To curList.count
            Set curItem = curList(i)
            Set listX = .ListItems.add(, , curItem.Id)
            For j = 1 To FormWorkMaterialType.fwmtMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormWorkMaterialType.fwmtMaterialName).text = _
                                            curItem.Name
            listX.ListSubItems(FormWorkMaterialType.fwmtIsTransport).text = _
                                            BoolToYesNo(curItem.IsTransport)
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    GoTo cleanHandler

errHandler:
    MsgBox Err.Number & vbCr & Err.Source & vbCr & Err.Description, _
                                                        vbCritical, "������"
    
cleanHandler:
    Set curItem = Nothing
    Set listX = Nothing
End Sub


Private Sub reloadFinanceSources()
'-----------------------------------------------------------------------------
' ���������� ������ ���������� ��������������
' Last update: 29.05.2019
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curItem As fsource
    
    
    With Me.ListViewList
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormFSourceEnum.ffsMax + 1
                .add
            Next i
            .Item(FormFSourceEnum.ffsId + 1).text = "���"
            .Item(FormFSourceEnum.ffsFromSubaccount + 1).text = _
                                                        "������ � ��������"
            .Item(FormFSourceEnum.ffsName + 1).text = "��������"
            .Item(FormFSourceEnum.ffsNote + 1).text = "����������"
        End With
        .ColumnHeaders(FormFSourceEnum.ffsId + 1).Width = 0
        
        ' ���������� �������
        .ListItems.Clear
        fsources.reload
        For i = 1 To fsources.count
            Set curItem = fsources(i)
            Set listX = .ListItems.add(, , curItem.Id)
            For j = 1 To FormFSourceEnum.ffsMax
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormFSourceEnum.ffsFromSubaccount).text = _
                                            BoolToYesNo(curItem.FromSubaccount)
            listX.ListSubItems(FormFSourceEnum.ffsName).text = curItem.Name
            listX.ListSubItems(FormFSourceEnum.ffsNote).text = curItem.Note
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set curItem = Nothing
    Set listX = Nothing
End Sub


Private Sub reloadMunicipalDistrict()
'-----------------------------------------------------------------------------
' ���������� ������ ������������� �����������
' Last update: 06.11.2019
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curLItem As address_md_class
    
    With Me.ListViewList
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormMDEnum.fmdMax + 1
                .add
            Next i
            .Item(FormMDEnum.fmdHead + 1).text = "�����"
            .Item(FormMDEnum.fmdHeadPosition + 1).text = _
                                                        "��������� �����"
            .Item(FormMDEnum.fmdId + 1).text = "���"
            .Item(FormMDEnum.fmdName + 1).text = "��������"
        End With
        .ColumnHeaders(FormMDEnum.fmdId + 1).Width = 0
        
        ' ���������� �������
        .ListItems.Clear
        address_md_list.reload
        For i = 1 To address_md_list.count
            Set curLItem = address_md_list(i)
            Set listX = .ListItems.add(, , curLItem.Id)
            For j = 1 To FormMDEnum.fmdMax
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormMDEnum.fmdHead).text = curLItem.Head
            listX.ListSubItems(FormMDEnum.fmdName).text = curLItem.Name
            listX.ListSubItems(FormMDEnum.fmdHeadPosition).text = _
                                                    curLItem.HeadPosition
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set curLItem = Nothing
    Set curItem = Nothing
    Set listX = Nothing
End Sub


Private Sub reloadHCounterPartTypes()
'-----------------------------------------------------------------------------
' ���������� ������ ������ ���� ��������������
' Last update: 02.07.2020
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curLItem As hcounter_part_type
    
    With Me.ListViewList
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormHCounterPartType.fhcptMax + 1
                .add
            Next i
            .Item(FormHCounterPartType.fhcptId + 1).text = "���"
            .Item(FormHCounterPartType.fhcptName + 1).text = "��������"
        End With
        .ColumnHeaders(FormHCounterPartType.fhcptId + 1).Width = 0
        
        ' ���������� �������
        .ListItems.Clear
        hcounter_part_types.reload
        For i = 1 To hcounter_part_types.count
            Set curLItem = hcounter_part_types(i)
            Set listX = .ListItems.add(, , curLItem.Id)
            For j = 1 To FormHCounterPartType.fhcptMax
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormHCounterPartType.fhcptId).text = curLItem.Name
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set curLItem = Nothing
    Set listX = Nothing
End Sub


Private Sub reloadTmpCounters()
'-----------------------------------------------------------------------------
' ���������� ������ ����� ������� ����
' Last update: 04.09.2020
'-----------------------------------------------------------------------------
    Dim curList As New tmp_counters
    curList.initialAll
    curList.fillListView Me.ListViewList
    Set curList = Nothing
    Set curItem = Nothing
End Sub


Private Sub gotoTmpCounterBuilding()
'-----------------------------------------------------------------------------
' �������� ����� ���� � ��������� ����
' Last update: 08.09.2020
'-----------------------------------------------------------------------------
    If Not curItem Is Nothing Then
        Dim BldnId As Long
        BldnId = curItem.ListSubItems(FormTmpCounters.ftcBldnId)
        Unload Me
        Call RunBuildingForm(BldnId)
    End If
End Sub


Private Function SetCurObject() As basicIdNameClass
'-----------------------------------------------------------------------------
' ��������� ������� �������
' 23.03.2022
'-----------------------------------------------------------------------------
    Select Case formType
        Case ListFormType.lftContractors
            Set SetCurObject = New contractor_class
        Case ListFormType.lftCounterModels
            Set SetCurObject = New counter_model
        Case ListFormType.lftExpenseItems
            Set SetCurObject = New expense_item
        Case ListFormType.lftFinanceSources
            Set SetCurObject = New fsource
        Case ListFormType.lftWorkMaterialTypes
            Set SetCurObject = New material_type
        Case ListFormType.lftMunicipalDistrict
            Set SetCurObject = New address_md_class
    End Select
End Function


Private Function SetCurObject1() As base_form_class
'-----------------------------------------------------------------------------
' ��������� ������� ������� ������ base_form_class
' 19.10.2022
'-----------------------------------------------------------------------------
    Select Case formType
        Case ListFormType.lftExpenseGroups
            Set SetCurObject1 = New expense_group
        Case ListFormType.lftExpenseItems
            Set SetCurObject1 = New expense_item
        Case ListFormType.lftCommonPropertyGroup
            Set SetCurObject1 = New common_property_group
        Case ListFormType.lftCommonPropertyElement
            Set SetCurObject1 = New common_property_element
        Case ListFormType.lftCommonPropertyParameter
            Set SetCurObject1 = New common_property_parameter
        Case ListFormType.lftManHourCostModes
            Set SetCurObject1 = New man_hour_cost_mode
        Case ListFormType.lftManHourCost
            Set SetCurObject1 = New man_hour_cost
        Case ListFormType.lftChairmansSign
            Set SetCurObject1 = New bldn_chairman_sign
    End Select
End Function


Private Sub ExportCurrentList()
'-----------------------------------------------------------------------------
' ����� � Excel �������� ���������� �����
' 08.10.2021
'-----------------------------------------------------------------------------
    Dim listCounter As Integer, itemCounter As Integer
    Dim ws As Worksheet, curRow As Integer, curColumn As Integer
    Dim curOut As Object
    
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    
    Me.Hide
    
    Set ws = ThisWorkbook.Worksheets.add
    curRow = 1
    ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, Me.ListViewList.ColumnHeaders.count)).Merge
    ws.Cells(curRow, 1).Value = Me.Caption
    curRow = curRow + 1
    With Me.ListViewList
        curColumn = 0
        For itemCounter = 1 To .ColumnHeaders.count
            curColumn = curColumn + 1
            ws.Cells(curRow, curColumn) = .ColumnHeaders(itemCounter).text
        Next itemCounter
        For listCounter = 1 To .ListItems.count
            curRow = curRow + 1
            curColumn = 1
            ws.Cells(curRow, curColumn) = .ListItems(listCounter).text
            For itemCounter = 1 To .ListItems(listCounter).ListSubItems.count
                curColumn = curColumn + 1
                ws.Cells(curRow, curColumn) = .ListItems(listCounter).ListSubItems(itemCounter).text
            Next itemCounter
        Next listCounter
    End With
    
    ws.Move
    Unload Me
    GoTo cleanHandler

errHandler:
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.delete
        Application.DisplayAlerts = True
    End If
    
cleanHandler:
    Application.ScreenUpdating = SUStatus
End Sub
