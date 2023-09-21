Attribute VB_Name = "RunForms"
Option Explicit

Sub RunAdmServiceForm()
' ----------------------------------------------------------------------------
' ������ ����� ������
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    AdmServiceForm.show
End Sub


Sub RunIdentificationForm()
' ----------------------------------------------------------------------------
' ������ ����� ��������������
' Last update: 26.10.2016
' ----------------------------------------------------------------------------
    IdentificationForm.show
End Sub


Sub RunReportSelectForm()
' ----------------------------------------------------------------------------
' ������ ����� ������ ������
' Last update: 01.06.2016
' ----------------------------------------------------------------------------
    SelectionForm.Tag = "report"
    SelectionForm.show
End Sub


Sub RunDictionarySelectForm()
' ----------------------------------------------------------------------------
' ������ ����� ������ �����������
' Last update: 07.06.2016
' ----------------------------------------------------------------------------
    SelectionForm.Tag = "dictionary"
    SelectionForm.show
End Sub


Sub RunImprovementForm()
' ----------------------------------------------------------------------------
' ������ ����� �� ��������� ���������������
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    ImprovementForm.show
End Sub


Sub RunVillageForm()
' ----------------------------------------------------------------------------
' ������ ����� � ���������� ��������
' Last update: 24.02.2016
' ----------------------------------------------------------------------------
    VillageForm.show
End Sub


Sub RunStreetForm()
' ----------------------------------------------------------------------------
' ������ ����� � �������
' Last update: 12.04.2016
' ----------------------------------------------------------------------------
    StreetForm.show
End Sub


Sub ShowWorkForm(BldnId As Long, prefWork As Long, _
                                    Optional fsMode As Integer = vbModal)
' ----------------------------------------------------------------------------
' ������ ����� ��������/��������� �����
' Last update: 15.11.2016
' ----------------------------------------------------------------------------
    WorkForm.BldnId = BldnId
    WorkForm.formPrefGWT = prefWork
    WorkForm.show fsMode
End Sub


Sub RunUKForm()
' ----------------------------------------------------------------------------
' ������ ����� ����������� ��������
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    UKForm.show
End Sub


Sub RunEmployeeForm(org As uk_class)
' ----------------------------------------------------------------------------
' ������ ����� �����������
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Set EmployeeForm.curMC = org
    EmployeeForm.show
End Sub


Sub RunGlobalWorkTypeForm()
' ----------------------------------------------------------------------------
' ������ ����� ���������� ����� �����
' Last update: 12.04.2016
' ----------------------------------------------------------------------------
    GlobalWorkForm.show
End Sub


Sub RunWorkTypeForm()
' ----------------------------------------------------------------------------
' ������ ����� ����� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    IDNameForm.objectTypeId = edfWorkType
    IDNameForm.show
End Sub


Sub RunWorkKindForm()
' ----------------------------------------------------------------------------
' ������ ����� ����� �����
' Last update: 29.03.2016
' ----------------------------------------------------------------------------
    WorkKindForm.show
End Sub


Sub RunBuildingForm(Optional BldnId As Long = NOTVALUE)
' ----------------------------------------------------------------------------
' ������ ����� ���� (� ������������ �������� ��� ����)
' Last update: 28.09.2016
' ----------------------------------------------------------------------------
    Load BuildingForm
    BuildingForm.formBldnId = BldnId
    BuildingForm.show ' vbModeless
End Sub


Sub RunBldnSelectionForm(reportType As BldnListReportType)
' ----------------------------------------------------------------------------
' ������ ����� ������ ������ �����
' 27.09.2022
' ----------------------------------------------------------------------------
    BldnReportSelectForm.reportType = reportType
    If (getPrefetchWork = SERVICE_GLOBAL_TYPE And reportType = blrPassport) Or _
            reportType = blrWorkCompletition Then
        BldnReportSelectForm.show
        Exit Sub
    End If
End Sub


Sub RunReport1Form(Optional formTag As String = "")
' ----------------------------------------------------------------------------
' ������ ����� ������ ���������� ������ 1
' Last update: 30.06.2016
' ----------------------------------------------------------------------------
    Report1Form.Tag = formTag
    Report1Form.show
End Sub


Sub RunAddBldnForm()
' ----------------------------------------------------------------------------
' ������ ����� ���������� ����
' Last update: 07.12.2016
' ----------------------------------------------------------------------------
    AddBldnForm.show vbModeless
End Sub


Sub RunWorkGroupInput()
' ----------------------------------------------------------------------------
' ������ ����� ����� ����� �� ������ �����
' Last update: 15.06.2017
' ----------------------------------------------------------------------------
    GroupWorkInputForm.show
End Sub


Sub RunWMForm()
' ----------------------------------------------------------------------------
' ������ ����� ���������� ����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    IDNameForm.objectTypeId = edfWallMaterial
    IDNameForm.show
End Sub


Sub RunDogovorForm()
' ----------------------------------------------------------------------------
' ������ ����� ����� ���������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    DogovorForm.show
End Sub


Sub RunBldnExpenseNameForm()
' ----------------------------------------------------------------------------
' ������ ����� �������� ������ �������� � �����
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    ExpenseNamesForm.show
End Sub


Sub RunServiceForm()
' ----------------------------------------------------------------------------
' ������ ����� �����
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    ServicesForm.show vbModeless
End Sub


Sub RunUserListForm()
' ----------------------------------------------------------------------------
' ������ ����� �������������
' Last update: 19.08.2018
' ----------------------------------------------------------------------------
    UserListForm.show
End Sub


Sub RunUserRolesForm()
' ----------------------------------------------------------------------------
' ������ ����� ����� �������������
' Last update: 19.08.2018
' ----------------------------------------------------------------------------
    UserRolesForm.show
End Sub


Sub RunAdminForm()
' ----------------------------------------------------------------------------
' ������ ����� �����������������
' Last update: 19.08.2018
' ----------------------------------------------------------------------------
    If CurrentUser.isAdmin Then AdminForm.show
End Sub


Sub RunUserRolesAccessForm()
' ----------------------------------------------------------------------------
' ������ ����� ������� �����
' Last update: 28.08.2018
' ----------------------------------------------------------------------------
    UserRolesAccessForm.show
End Sub


Sub RunTmpCounterForm(Address As String, BldnId As Long, _
                        Optional ItemId As Long = NOTVALUE, _
                        Optional isChange As Boolean = False)
' ----------------------------------------------------------------------------
' ������ ����� ��������
' Last update: 15.09.2020
' ----------------------------------------------------------------------------
    With TmpCountersForm
        .Address = Address
        .BldnId = BldnId
        .isChange = isChange
        If isChange Then
            .Item = ItemId
        End If
        .show
    End With
End Sub


Sub RunListForm(formType As ListFormType)
' ----------------------------------------------------------------------------
' ������ ����� ListForm � ������ ����������
' 21.09.2021
' ----------------------------------------------------------------------------
    With ListForm
        .formType = formType
        .show
    End With
End Sub