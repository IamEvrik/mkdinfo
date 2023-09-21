VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Report1Form 
   Caption         =   "����� ���������� ������"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   OleObjectBlob   =   "Report1Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Report1Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BEGIN_TERM_TAG As String = "bTerm"

Private Const BEGIN_TERM_LABEL As String = "��������� ������"
Private Const END_TERM_LABEL As String = "�������� ������"
Private Const MD_LABEL As String = "������������� �����������"
Private Const CONTRACTOR_LABEL As String = "��������� �����������"
Private Const MC_LABEL As String = "����������� ��������"
Private Const DOGOVOR_LABEL As String = "��� ��������"
Private Const GWT_LABEL As String = "��� �������"
Private Const WT_LABEL As String = "��� ������"
Private Const WK_LABEL As String = "��� ������"
Private Const ALL_WORKS_LABEL As String = "�������� �������� ������"
Private Const UK_SERVICE_LABEL As String = "������"
Private Const ADDED_TYPE_LABEL As String = "��� �������"
Private Const VILLAGE_LABEL As String = "���������� �����"
Private Const ADDRESS_LABEL As String = "�����"

' �������, ��������� ��� ��� �������� ������
Private updateEndTerm As Boolean

Private updateCombo2 As Boolean


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� ����� - ������� �������� ��� ������� ������ ���������
' 07.02.2023
' ----------------------------------------------------------------------------
    Dim formTag As String
    
    formTag = Split(Me.Tag, ":")(0)
    
    Call hideAllElements
    
    Select Case LCase(formTag)
        Case "report_1":                Call paintFormReport1
        Case "report_2":                Call paintFormReport2
        Case "report_3":                Call paintFormReport2
        Case "report_4":                Call paintFormReport4
        Case "report_7":                Call paintFormWorkTypes
        Case "report_9":                Call paintFormReport9
        Case "works":                   Call paintFormWorks
        Case "money":                   Call paintFormMoney
        Case "passport", "tosite":      Call paintFormBldnPassport
        Case "allworks":                Call paintFormAllWorks
        Case "subaccount":              Call paintFormForSubaccounts
        Case "bldnwork":                Call paintFormBldnWork
        Case "gisexpense":              Call paintFormGisExpense
        Case "report_mworkmaterials":   Call paintFormMWorkMaterials
        Case "report_contmaterials":    Call paintFormMWorkMaterials
        Case "report_10":               Call paintFormReport10
        Case "report_101":              Call paintFormReport101
        Case "report_101a":             Call paintFormReport101a
        Case "report_102":              Call paintFormReport102
        Case "report_11":               Call paintFormReport11
        Case "report_12":               Call paintFormReport12
        Case "report_201":              Call paintFormReport201
        Case "report_13":              Call paintFormReport13
        Case "report_14":              Call paintFormReport14
    End Select
End Sub


Private Sub ComboBox1_Change()
' ----------------------------------------------------------------------------
' ���������� ����� ��� ��������� ��� �������������
' 22.11.2021
' ----------------------------------------------------------------------------
    If StrComp(Me.ComboBox1.Tag, BEGIN_TERM_TAG, vbTextCompare) = 0 And _
                                                        updateEndTerm Then
        Call loadEndDate(1)
    ElseIf StrComp(Me.Label1, MD_LABEL, vbTextCompare) = 0 And _
            StrComp(Me.Tag, "report_12", vbTextCompare) = 0 And _
            updateCombo2 Then
        Call reloadComboBox(rcmListManagedBldnAddressIdByMD, Me.ComboBox2, initValue:=Me.ComboBox1)
        Me.ComboBox2.ListIndex = 0
    End If
End Sub


Private Sub ComboBox2_Change()
' ----------------------------------------------------------------------------
' ��� ��������� �������� ����� ��������� ��������� ������
' 17.09.2021
' ----------------------------------------------------------------------------
    If StrComp(Me.ComboBox2.Tag, BEGIN_TERM_TAG, vbTextCompare) = 0 And _
                                                        updateEndTerm Then
        Call loadEndDate(2, False)
    ElseIf StrComp(Me.Label3, VILLAGE_LABEL, vbTextCompare) = 0 Then
        Call reloadComboBox(rcmVillage, Me.ComboBox3, _
                initValue:=Me.ComboBox2, addAllItems:=True)
    End If
End Sub


Private Sub combobox3_Change()
' ----------------------------------------------------------------------------
' ��� ��������� �������� ����� ��������� ��������� ������
' 17.09.2021
' ----------------------------------------------------------------------------
    If StrComp(Me.ComboBox3.Tag, BEGIN_TERM_TAG, vbTextCompare) = 0 And _
                                                        updateEndTerm Then
        Call loadEndDate(3)
    ElseIf StrComp(Me.Label4, VILLAGE_LABEL, vbTextCompare) = 0 Then
        Call reloadComboBox(rcmVillage, Me.ComboBox4, _
                initValue:=Me.ComboBox3, addAllItems:=True)
    End If
End Sub


Private Sub ComboBox6_Change()
' ----------------------------------------------------------------------------
' ���������� �����
' Last update: 19.09.2016
' ----------------------------------------------------------------------------
    If Me.ComboBox6.ListIndex > -1 Then
        If Me.ComboBox6.Value <> ALLVALUES Then
            Call reloadComboBox(rcmWorkKind, Me.ComboBox7, _
                    initValue:=Me.ComboBox6.Value, addAllItems:=True)
        Else
            Call reloadComboBox(rcmWorkKind, Me.ComboBox7, _
                    addAllItems:=True)
        End If
    End If
End Sub


Private Sub ComboBox8_Change()
' ----------------------------------------------------------------------------
' ��� ��������� ��������� ���� ��������������� ��������
' Last update: 19.09.2016
' ----------------------------------------------------------------------------
    If StrComp(Me.ComboBox8.Tag, BEGIN_TERM_TAG, vbTextCompare) = 0 And _
                                                        updateEndTerm Then
        Call loadEndDate(8)
    End If
End Sub


Private Sub ButtonCreateReport_Click()
' ----------------------------------------------------------------------------
' ������� ������ ������������ �������
' 20.12.2022
' ----------------------------------------------------------------------------
    Dim wtCol As Collection
    Dim i As Long
    Dim formTag As String
    Dim ItemId As Long
    Dim fullTag As String
    
    fullTag = Me.Tag
    formTag = Split(Me.Tag, ":")(0)
    If InStr(1, Me.Tag, ":") > 0 Then
        ItemId = Split(Me.Tag, ":")(1)
    End If
    Me.Hide ' ��� ����� Excel 2016 �� ������ �������� ��������� �����
    Unload Me
    
    Select Case LCase(formTag)
        Case "report_2"
            Call Report_2( _
                        mcId:=Me.ComboBox1.Value, _
                        dogovorId:=Me.ComboBox2.Value, _
                        mdId:=Me.ComboBox3.Value, _
                        contractorId:=Me.ComboBox4.Value)
        Case "report_3"
            Call Report_3( _
                        mcId:=Me.ComboBox1.Value, _
                        dogovorId:=Me.ComboBox2.Value, _
                        mdId:=Me.ComboBox3.Value, _
                        contractorId:=Me.ComboBox4.Value)
        Case "report_4"
            Call Report_4( _
                beginMonth:=Me.ComboBox8.list(Me.ComboBox8.ListIndex, ccname), _
                endMonth:=Me.ComboBox9.list(Me.ComboBox9.ListIndex, ccname), _
                contId:=Me.ComboBox4.Value, _
                mcId:=Me.ComboBox1.Value, _
                mdId:=Me.ComboBox3.Value, _
                gwtId:=Me.ComboBox5.Value, _
                wtId:=Me.ComboBox6.Value, _
                wkId:=Me.ComboBox7.Value, _
                Status:=Me.ComboBox10.Value, _
                Dogovor:=Me.ComboBox2.Value)
        Case "money"
            Call MoneyReport(Me.ComboBox3.Value, _
                        Me.ComboBox4.Value, _
                        Me.ComboBox2.Value, _
                        Me.ComboBox1.Value)
        Case "works"
            Call WorkReport(Me.ComboBox3.Value, _
                        Me.ComboBox4.Value, _
                        Me.ComboBox2.Value, _
                        Me.ComboBox1.Value)
        Case "passport"
            Call BldnPassport(ItemId:=ItemId, _
                            not_show_sum:=Not CBool(Me.ComboBox3.Value), _
                            beginDate:=Me.ComboBox1.Value, _
                            EndDate:=Me.ComboBox2.Value)
        Case "report_1"
            Call Report_1( _
                beginMonth:=Me.ComboBox8.Value, _
                endMonth:=Me.ComboBox9.Value, _
                contId:=Me.ComboBox4.Value, _
                mcId:=Me.ComboBox1.Value, _
                mdId:=Me.ComboBox3.Value, _
                gwtId:=Me.ComboBox5.Value, _
                wtId:=Me.ComboBox6.Value, _
                wkId:=Me.ComboBox7.Value, _
                bMonthName:=Me.ComboBox8.list(Me.ComboBox8.ListIndex, ccname), _
                eMonthName:=Me.ComboBox9.list(Me.ComboBox9.ListIndex, ccname), _
                needLess:=Me.ComboBox10, _
                dogovorId:=Me.ComboBox2.Value)
        Case "allworks"
            Call AllWorksReport(Me.ComboBox1.Value)
        Case "subaccount"
            Call SubAccountReport(beginMonth:=Me.ComboBox2.Value, _
                                endMonth:=Me.ComboBox3.Value, _
                                gwtId:=Me.ComboBox1.Value)
        Case "tosite"
            Call reportToSite(beginDate:=Me.ComboBox1.Value, _
                            EndDate:=Me.ComboBox2.Value, _
                            reportYear:=Year(terms(CStr( _
                                Me.ComboBox2.Value)).beginDate), _
                            not_show_sum:=Not CBool(Me.ComboBox3.Value))
        Case "report_7"
            Call Report_7(Me.ComboBox1.Value)
        Case "report_9"
            Call Report_9( _
                beginMonth:=Me.ComboBox8.Value, _
                endMonth:=Me.ComboBox9.Value, _
                contId:=Me.ComboBox4.Value, _
                mcId:=Me.ComboBox1.Value, _
                mdId:=Me.ComboBox3.Value, _
                gwtId:=Me.ComboBox5.Value, _
                wtId:=Me.ComboBox6.Value, _
                wkId:=Me.ComboBox7.Value, _
                bMonthName:=Me.ComboBox8.list(Me.ComboBox8.ListIndex, ccname), _
                eMonthName:=Me.ComboBox9.list(Me.ComboBox9.ListIndex, ccname), _
                fSourceId:=Me.ComboBox10, _
                dogovorId:=Me.ComboBox2.Value)
        Case "bldnwork"
            Call ReportBldnWorks(CLng(Split(fullTag, ":")(1)), _
                            Me.ComboBox1.Value, _
                            Me.ComboBox2.Value, _
                            Me.ComboBox3.Value, _
                            Me.ComboBox4.Value, _
                            Me.ComboBox5.Value)
        Case "gisexpense"
            Dim mCount As Long
            mCount = DateDiff("m", dateValue(Me.ComboBox1.text), _
                                        dateValue(Me.ComboBox2.text))
            Call reportBldnPlanExpenseToGis(CLng(Split(fullTag, ":")(1)), _
                                            Me.ComboBox1.Value, mCount + 1)
        Case "report_mworkmaterials"
            Call Report_MWorkMaterials( _
                    Contractor:=Me.ComboBox1.Value, _
                    beginDate:=Me.ComboBox2.Value, _
                    EndDate:=Me.ComboBox3.Value)
        Case "report_contmaterials"
            Call Report_ContractorMaterials( _
                    Contractor:=Me.ComboBox1.Value, _
                    beginDate:=Me.ComboBox2.Value, _
                    EndDate:=Me.ComboBox3.Value)
        Case "report_10"
            Call ReportSubAccountsPlan(InContractorId:=Me.ComboBox1.Value)
        Case "report_101"
            Call report_101( _
                InUkServiceId:=Me.ComboBox1, _
                InTermId:=Me.ComboBox2)
        Case "report_101a"
            Call report_101a( _
                InUkServiceId:=Me.ComboBox1, _
                InBeginTermId:=Me.ComboBox2, _
                InEndTermId:=Me.ComboBox3)
        Case "report_102"
            Call report_102( _
                InTypeId:=Me.ComboBox1, _
                InBeginTerm:=Me.ComboBox2, _
                InEndTerm:=Me.ComboBox3, _
                InTypeName:=Me.ComboBox1.text)
        Case "report_11"
            Call report_11( _
                InMCId:=Me.ComboBox2, _
                InMDId:=Me.ComboBox4, _
                InVillageId:=Me.ComboBox5, _
                InContractorId:=Me.ComboBox3, _
                InDate:=Me.ComboBox1, _
                InMCName:=Me.ComboBox2.text, _
                InMDName:=Me.ComboBox4.text, _
                InVillageName:=Me.ComboBox5.text, _
                InContractorName:=Me.ComboBox3.text, _
                InIsFull:=CBool(Me.ComboBox6.Value))
        Case "report_12"
            Call report_12( _
                InBldnId:=Me.ComboBox2, _
                InBeginTerm:=Me.ComboBox3, _
                InEndTerm:=Me.ComboBox4, _
                InAddress:=Me.ComboBox2.text, _
                InBeginDate:=Me.ComboBox3.text, _
                InEndDate:=Me.ComboBox4.text)
        Case "report_201"
            Call report_201( _
                InBeginTerm:=Me.ComboBox1, _
                InEndTerm:=Me.ComboBox2, _
                InGwtId:=Me.ComboBox3)
        Case "report_13"
            Call report_13( _
                InBeginTerm:=Me.ComboBox1, _
                InEndTerm:=Me.ComboBox2)
        Case "report_14"
            Call report_14( _
                InBeginTerm:=Me.ComboBox1, _
                InEndTerm:=Me.ComboBox2, _
                InUkService:=Me.ComboBox3)
        End Select
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' ������� ������ ������
' Last update: 20.09.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub paintFormReport1()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 1 (����� �� �������� �������)
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, MC_LABEL)
    Call reloadComboBox(rcmMC, Me.ComboBox1, addAllItems:=True)
    
    Call showHideElements("2", True, DOGOVOR_LABEL)
    Call reloadComboBox(rcmDogovor, Me.ComboBox2, addAllItems:=True)
    
    Call showHideElements("3", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox3, addAllItems:=True)
    
    Call showHideElements("4", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmContractor, Me.ComboBox4, addAllItems:=True)
    
    Call showHideElements("5", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox5, addAllItems:=True, _
                                                defValue:=getPrefetchWork())
                                                
    Call showHideElements("6", True, WT_LABEL)
    Call reloadComboBox(rcmWorkType, Me.ComboBox6, addAllItems:=True)
                                                
    Call showHideElements("7", True, WK_LABEL)
    Call reloadComboBox(rcmWorkKind, Me.ComboBox7, addAllItems:=True)
                                                
    Call showHideElements("8", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox8, _
                                        defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox8.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("9", True, END_TERM_LABEL)
    Call loadEndDate(8)
    
    Call showHideElements("10", True, ALL_WORKS_LABEL)
    Call reloadComboBox(rcmYesNo, Me.ComboBox10, defValue:=CLng(False), _
                                                initString:="������ ��������")
End Sub


Private Sub paintFormReport2()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 2 (����������� ����������)
' Last update: 11.04.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, MC_LABEL)
    Call reloadComboBox(rcmMC, Me.ComboBox1, addAllItems:=True)
    
    Call showHideElements("2", True, DOGOVOR_LABEL)
    Call reloadComboBox(rcmDogovor, Me.ComboBox2, addAllItems:=True)
    
    Call showHideElements("3", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox3, addAllItems:=True)
    
    Call showHideElements("4", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmMainContractor, Me.ComboBox4, _
                                                        addAllItems:=True)
End Sub


Private Sub paintFormReport4()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 4 (����� � �������� �������)
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, MC_LABEL)
    Call reloadComboBox(rcmMC, Me.ComboBox1, addAllItems:=True)
    
    Call showHideElements("2", True, DOGOVOR_LABEL)
    Call reloadComboBox(rcmDogovor, Me.ComboBox2, addAllItems:=True)
    
    Call showHideElements("3", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox3, addAllItems:=True)
    
    Call showHideElements("4", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmContractor, Me.ComboBox4, addAllItems:=True)
    
    Call showHideElements("5", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox5, addAllItems:=True, _
                                                defValue:=getPrefetchWork())
                                                
    Call showHideElements("6", True, WT_LABEL)
    Call reloadComboBox(rcmWorkType, Me.ComboBox6, addAllItems:=True)
                                                
    Call showHideElements("7", True, WK_LABEL)
    Call reloadComboBox(rcmWorkKind, Me.ComboBox7, addAllItems:=True)
                                                
    Call showHideElements("8", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmPlanTerms, Me.ComboBox8, _
                                initValue:=DateSerial(FIRST_PLAN_YEAR, 1, 1), _
                                initValue2:=Year(Now) - FIRST_PLAN_YEAR + 4, _
                                defValue:=DateSerial(Year(Now), 1, 1))
                                
    Call showHideElements("9", True, END_TERM_LABEL)
    Call reloadComboBox(rcmPlanTerms, Me.ComboBox9, _
                                initValue:=DateSerial(FIRST_PLAN_YEAR, 1, 1), _
                                initValue2:=Year(Now) - FIRST_PLAN_YEAR + 4, _
                                defValue:=DateSerial(Year(Now), 12, 1))
                                
    Call showHideElements("10", True, "������")
    Call reloadComboBox(rcmPlanStatuses, Me.ComboBox10, addAllItems:=True)
End Sub


Private Sub paintFormReport9()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 9 (����� �� �������� ������� � �����������
'                           ��������������)
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, MC_LABEL)
    Call reloadComboBox(rcmMC, Me.ComboBox1, addAllItems:=True)
    
    Call showHideElements("2", True, DOGOVOR_LABEL)
    Call reloadComboBox(rcmDogovor, Me.ComboBox2, addAllItems:=True)
    
    Call showHideElements("3", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox3, addAllItems:=True)
    
    Call showHideElements("4", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmContractor, Me.ComboBox4, addAllItems:=True)
    
    Call showHideElements("5", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox5, addAllItems:=True, _
                                                defValue:=getPrefetchWork())
                                                
    Call showHideElements("6", True, WT_LABEL)
    Call reloadComboBox(rcmWorkType, Me.ComboBox6, addAllItems:=True)
                                                
    Call showHideElements("7", True, WK_LABEL)
    Call reloadComboBox(rcmWorkKind, Me.ComboBox7, addAllItems:=True)
                                                
    Call showHideElements("8", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox8, _
                                        defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox8.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("9", True, END_TERM_LABEL)
    Call loadEndDate(8)
    
    Call showHideElements("10", True, "�������� ��������������")
    Call reloadComboBox(rcmFSources, Me.ComboBox10, addAllItems:=True)
End Sub


Private Sub paintFormMoney()
' ----------------------------------------------------------------------------
' ��������� ����� ������ �� �����������
' Last update: 04.02.2020
' ----------------------------------------------------------------------------
    updateEndTerm = False
    
    Call showHideElements("1", True, DOGOVOR_LABEL)
    Call reloadComboBox(rcmDogovor, Me.ComboBox1, addAllItems:=True)
    
    Call showHideElements("2", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmMainContractor, Me.ComboBox2, addAllItems:=True)
    
    Call showHideElements("3", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox3, _
                                        defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox3.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("4", True, END_TERM_LABEL)
    Call loadEndDate(3)
End Sub


Private Sub paintFormWorks()
' ----------------------------------------------------------------------------
' ��������� ����� ������ �� �������
' Last update: 03.10.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox1, defValue:=getPrefetchWork)
    
    Call showHideElements("2", True, CONTRACTOR_LABEL)
    Call reloadComboBox(IIf(getPrefetchWork = SERVICE_GLOBAL_TYPE, _
                            rcmMainContractor, rcmContractor), _
                        Me.ComboBox2, addAllItems:=True)
    
    updateEndTerm = False
    Call showHideElements("3", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox3, _
                                    defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox3.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("4", True, END_TERM_LABEL)
    Call loadEndDate(3)
End Sub


Private Sub paintFormBldnPassport()
' ----------------------------------------------------------------------------
' ��������� ����� �������� ����
' Last update: 15.04.2019
' ----------------------------------------------------------------------------
    updateEndTerm = False
    Call showHideElements("1", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox1, _
                                    defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox1.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("2", True, END_TERM_LABEL)
    Call loadEndDate(1)
    
    Call showHideElements("3", True, "�������� �����")
    Call reloadComboBox(rcmYesNo, Me.ComboBox3, defValue:=CInt(False))
End Sub


Private Sub paintFormAllWorks()
' ----------------------------------------------------------------------------
' ��������� ����� "��� ������"
' Last update: 15.04.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox1, addAllItems:=True)
End Sub


Private Sub paintFormForSubaccounts()
' ----------------------------------------------------------------------------
' ��������� ����� ������ ��� ���������
' Last update: 15.04.2019
' ----------------------------------------------------------------------------
    updateEndTerm = False
    
    Call showHideElements("1", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox1, defValue:=getPrefetchWork)
    
    Call showHideElements("2", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox2, _
                                        defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox2.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("3", True, END_TERM_LABEL)
    Call loadEndDate(2)
End Sub


Private Sub paintFormBldnWork()
' ----------------------------------------------------------------------------
' ��������� ����� ������ �� ������� �� ����
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox1, addAllItems:=True, _
                                    defValue:=longValue(getPrefetchWork()))
    
    Call showHideElements("2", True, WT_LABEL)
    Call reloadComboBox(rcmWorkType, Me.ComboBox2, addAllItems:=True)
    
    updateEndTerm = False
    Call showHideElements("3", True, BEGIN_TERM_LABEL)
    Me.ComboBox3.Tag = BEGIN_TERM_TAG
    Call reloadComboBox(rcmTermDESC, Me.ComboBox3, addAllItems:=True, _
                                        defValue:=terms.FirstTermInYear.Id)
    updateEndTerm = True
    
    Call showHideElements("4", True, END_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox4, _
                        addAllItems:=True, defValue:=terms.LastTermInYear.Id)
    
    Call showHideElements("5", True, "�������� ��������������")
    Call reloadComboBox(rcmFSources, Me.ComboBox5, _
                                                            addAllItems:=True)
End Sub


Private Sub paintFormWorkTypes()
' ----------------------------------------------------------------------------
' ��������� ����� ������ � ������ �����
' Last update: 15.04.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, WT_LABEL)
    Call reloadComboBox(rcmWorkType, Me.ComboBox1, addAllItems:=True)
End Sub


Private Sub paintFormGisExpense()
' ----------------------------------------------------------------------------
' ��������� ����� ������ �������� ����� ����� � ���
' Last update: 17.05.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox1, _
                                        defValue:=terms.FirstTermInYear.Id)
    
    Call showHideElements("2", True, END_TERM_LABEL)
    Call reloadComboBox(rcmPlanTerms, Me.ComboBox2, _
                                initValue:=terms.FirstTermInYear.beginDate, _
                                initValue2:=3, defValue:=12)
End Sub


Private Sub paintFormMWorkMaterials()
' ----------------------------------------------------------------------------
' ��������� ����� ������ �� ���������� � �����������
' Last update: 31.10.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmMainContractor, Me.ComboBox1, addAllItems:=True)
    
    Call showHideElements("2", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox2, _
                                        defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox2.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("3", True, END_TERM_LABEL)
    Call loadEndDate(2, False)
End Sub


Private Sub paintFormReport10()
' ----------------------------------------------------------------------------
' ��������� ����� ������ ��� ������������ ���������
' Last update: 28.09.2020
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmMainContractor, Me.ComboBox1, addAllItems:=True)
End Sub


Private Sub paintFormReport101()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 101 (�������� ����������)
' 12.10.2021
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, UK_SERVICE_LABEL)
    Call reloadComboBox(rcmUkServices, Me.ComboBox1, addAllItems:=True)
    Me.ComboBox1.ListIndex = 0
    
    Call showHideElements("2", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox2)
    Me.ComboBox2.ListIndex = 0
End Sub


Private Sub paintFormReport101a()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 101a (�������� ���������� �� ������)
' 12.10.2021
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, UK_SERVICE_LABEL)
    Call reloadComboBox(rcmUkServices, Me.ComboBox1, addAllItems:=True)
    Me.ComboBox1.ListIndex = 0
    
    Call showHideElements("2", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox2)
    Me.ComboBox2.ListIndex = 0
    Me.ComboBox2.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("3", True, END_TERM_LABEL)
    Call loadEndDate(2, False)
End Sub


Private Sub paintFormReport102()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 102 (������� ���������� �� ������)
' 15.09.2021
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, ADDED_TYPE_LABEL)
    Call reloadComboBox(rcmAddedTypes, Me.ComboBox1)
    Me.ComboBox1.ListIndex = 0
    
    Call showHideElements("2", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox2)
    Me.ComboBox2.ListIndex = 0
    Me.ComboBox2.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("3", True, END_TERM_LABEL)
    Call loadEndDate(2, False)
End Sub


Private Sub paintFormReport11()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 11 (������� ��������� ���������)
' 11.05.2022
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmSubaccountTerms, Me.ComboBox1)
    Me.ComboBox1.ListIndex = 0
    
    Call showHideElements("2", True, MC_LABEL)
    Call reloadComboBox(rcmMC, Me.ComboBox2, addAllItems:=True)
    
    Call showHideElements("3", True, CONTRACTOR_LABEL)
    Call reloadComboBox(rcmUsingMainContractor, Me.ComboBox3, addAllItems:=True)
    
    Call showHideElements("4", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox4, addAllItems:=True)
    
    Call showHideElements("5", True, VILLAGE_LABEL)
    Call reloadComboBox(rcmVillage, Me.ComboBox5, initValue:=Me.ComboBox3.Value, addAllItems:=True)
    
    Call showHideElements("6", True, "������ ����������")
    Call reloadComboBox(rcmYesNo, Me.ComboBox6)
    Me.ComboBox6.ListIndex = 0
End Sub


Private Sub paintFormReport12()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 12 (������� ������������)
' 22.11.2021
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, MD_LABEL)
    Call reloadComboBox(rcmMd, Me.ComboBox1)
    Me.ComboBox1.ListIndex = 0
    
    Call showHideElements("2", True, ADDRESS_LABEL)
    Call reloadComboBox(rcmListManagedBldnAddressIdByMD, Me.ComboBox2, initValue:=Me.ComboBox1)
    Me.ComboBox2.ListIndex = 0
    updateCombo2 = True
    
    Call showHideElements("3", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox3, defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox3.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("4", True, END_TERM_LABEL)
    Call loadEndDate(3, False)
End Sub


Private Sub paintFormReport201()
' ----------------------------------------------------------------------------
' ��������� ����� ������ 201 (��� ����)
' 07.06.2022
' ----------------------------------------------------------------------------
    Call showHideElements("1", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTermDESC, Me.ComboBox1, defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox1.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("2", True, END_TERM_LABEL)
    Call loadEndDate(1, False)
    
    Call showHideElements("3", True, GWT_LABEL)
    Call reloadComboBox(rcmGWT, Me.ComboBox3, defValue:=getPrefetchWork())
    Me.ComboBox3.ListIndex = 0
End Sub


Private Sub paintFormReport13()
' ----------------------------------------------------------------------------
' ��������� ����� ����� 13 ������ ������� �����
' 20.12.2022
' ----------------------------------------------------------------------------
    updateEndTerm = False
    Call showHideElements("1", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox1, _
                                    defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox1.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("2", True, END_TERM_LABEL)
    Call loadEndDate(1)
End Sub


Private Sub paintFormReport14()
' ----------------------------------------------------------------------------
' ��������� ����� ����� 14 ������������ �� �����
' 07.02.2023
' ----------------------------------------------------------------------------
    updateEndTerm = False
    Call showHideElements("1", True, BEGIN_TERM_LABEL)
    Call reloadComboBox(rcmTerm, Me.ComboBox1, _
                                    defValue:=terms.FirstTermInYear.Id)
    Me.ComboBox1.Tag = BEGIN_TERM_TAG
    updateEndTerm = True
    
    Call showHideElements("2", True, END_TERM_LABEL)
    Call loadEndDate(1)
    
    Call showHideElements("3", True, UK_SERVICE_LABEL)
    Call reloadComboBox(rcmUkServices, Me.ComboBox3)
    Me.ComboBox3.ListIndex = 0
    
End Sub


Private Sub loadEndDate(curCombo As Long, _
                                    Optional addAllValues As Boolean = True)
' ----------------------------------------------------------------------------
' ���������� ������ ��������� ������� (��������� �������,
'        ������� ������ ���������� ����������)
' Last update: 31.10.2019
' ----------------------------------------------------------------------------
    Dim curDate As Date
    
    If Me.Controls("ComboBox" & curCombo).Value = ALLVALUES Then
        Me.Controls("ComboBox" & curCombo + 1).ListIndex = 0
    Else
        curDate = terms(CStr(Me.Controls("ComboBox" & curCombo).Value)).beginDate
        Call reloadComboBox(rcmTermDESC, _
                            Me.Controls("ComboBox" & curCombo + 1), _
                            addAllItems:=addAllValues, _
                            defValue:=terms.LastTermInYear(Year(curDate)).Id)
        With Me.Controls("ComboBox" & curCombo + 1)
            For i = .ListCount - 1 To 0 Step -1
                If .list(i, ccId) <> ALLVALUES Then
                    If terms(.list(i, ccId)).beginDate < curDate Then
                        .RemoveItem i
                    End If
                End If
            Next i
        End With
    End If
End Sub


Private Sub hideAllElements()
' ----------------------------------------------------------------------------
' ���������� ��� �������� �����
' Last update: 12.04.2019
' ----------------------------------------------------------------------------
    Call showHideElements("1", False)
    Call showHideElements("2", False)
    Call showHideElements("3", False)
    Call showHideElements("4", False)
    Call showHideElements("5", False)
    Call showHideElements("6", False)
    Call showHideElements("7", False)
    Call showHideElements("8", False)
    Call showHideElements("9", False)
    Call showHideElements("10", False)
End Sub


Private Sub showHideElements(elementNum As String, elementVisible As Boolean, _
                                        Optional elementText As String = "")
' ----------------------------------------------------------------------------
' ������/�������� ��������� ������� � ��������� ����������
' Last update: 12.04.2019
' ----------------------------------------------------------------------------
    Me.Controls("Label" & elementNum).Caption = elementText
    Me.Controls("Label" & elementNum).Visible = elementVisible
    Me.Controls("ComboBox" & elementNum).Visible = elementVisible
End Sub

