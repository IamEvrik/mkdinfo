VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "building_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ���
' ----------------------------------------------------------------------------

Private classId As Long                         ' ��� ����
Private classStreet As Long                     ' �����
Private classBldnNo As String                   ' ����� ����
Private classSiteNo As String                   ' ����� ��� �����
Private classHotWater As Long                   ' ������� ����
Private classHeating As Long                    ' ���������
Private classGas As Long                        ' ���
Private classColdWater As Long                  ' �������� ����
Private classUK As Long                         ' ����������� ��������
Private classContractor As Long                 ' ���������
Private classImprovement As Long                ' ������� ���������������
Private classManagement As Long                 ' ��� ��������
Private classOutReport As Boolean               ' �������� �� ����� �� ����
Private classCadastral As String                ' ����������� ����� ������
Private classContractNo As String               ' ����� ��������
Private classContractDate As Date               ' ���� ��������
Private classBldnType As Long                   ' ��� ����
Private classDisRepair As Boolean               ' �������� ���������
Private classEnergoClass As Long                ' ����� �������������������
Private classSubAccount As Currency             ' ������� �� ��������
Private classSubAccountDateID As Integer        ' ����, �� ������� �������
Private classFias As String                     ' ��� ����
Private classGisGuid As String                  ' ���� � ��� ���
Private classSubAccountPercent As Integer       ' ������� ������������
Private classPlanSubAccountSum As Currency      ' �������� �����������
Private classCurrentSubAccount As Currency      ' ������� �� �������� � ������ �����
Private classManHourCost As man_hour_cost       ' ����� ��������� ������������


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' Last update: 15.04.2016
' ----------------------------------------------------------------------------
    Call flushValues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����
' 26.05.2022
' ----------------------------------------------------------------------------
    If ItemId = NOTVALUE Then Exit Sub

    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "get_building"
    sqlParams.add "InItemId", ItemId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    classId = ItemId
    classStreet = rst!street_id
    classBldnNo = rst!bldn_no
    classUK = rst!mc_id
    classContractor = longValue(rst!contractor_id)
    classImprovement = longValue(rst!improvement_id)
    classOutReport = rst!Out_Report
    classHotWater = rst!hot_water
    classColdWater = rst!cold_water
    classHeating = rst!Heating
    classGas = rst!Gas
    classManagement = rst!dogovor_type
    classSiteNo = DBgetString(rst!site_no)
    classBldnType = longValue(rst!bldn_type)
    classCadastral = DBgetString(rst!cadastral_no)
    classContractNo = DBgetString(rst!contract_no)
    classContractDate = DBgetDate(rst!contract_date)
    classDisRepair = boolValue(rst!DisRepair)
    classEnergoClass = longValue(rst!energo_class)
    classFias = DBgetString(rst!Fias)
    classGisGuid = DBgetString(rst!gis_guid)
    rst.Close
    
    Call getManHourCost
    Call getSubAccount
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "building_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' ���
' Last update: 15.04.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Dogovor() As dogovor_class
' ----------------------------------------------------------------------------
' ��� �������
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Set Dogovor = dogovor_list(CStr(classManagement))
End Property


Property Get BldnNo() As String
' ----------------------------------------------------------------------------
' ����� ����
' Last update: 15.04.2016
' ----------------------------------------------------------------------------
    BldnNo = classBldnNo
End Property


Property Get SiteNo() As String
' ----------------------------------------------------------------------------
' ����� ���� ��� �����
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    SiteNo = classSiteNo
End Property


Property Get CadastralNo() As String
' ----------------------------------------------------------------------------
' ����������� ����� ������
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    CadastralNo = classCadastral
End Property


Property Get HotWater() As id_name_class
' ----------------------------------------------------------------------------
' ������� ������� ����
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmpList As New id_name_list
    tmpList.initHotWater
    Set HotWater = tmpList(CStr(classHotWater))
    Set tmpList = Nothing
End Property


Property Get Gas() As id_name_class
' ----------------------------------------------------------------------------
' ��� �������������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmpList As New id_name_list
    tmpList.initGas
    Set Gas = tmpList(CStr(classGas))
    Set tmpList = Nothing
End Property


Property Get Heating() As id_name_class
' ----------------------------------------------------------------------------
' ������� ���������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmpList As New id_name_list
    tmpList.initHeating
    Set Heating = tmpList(CStr(classHeating))
    Set tmpList = Nothing
End Property


Property Get ColdWater() As id_name_class
' ----------------------------------------------------------------------------
' �������� �������������
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    Dim tmpList As New id_name_list
    tmpList.initColdWater
    Set ColdWater = tmpList(CStr(classColdWater))
    Set tmpList = Nothing
End Property


Property Get Contractor() As contractor_class
' ----------------------------------------------------------------------------
' ���������
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Set Contractor = contractor_list(CStr(classContractor))
End Property


Property Get Improvement() As improvement_class
' ----------------------------------------------------------------------------
' ������� ���������������
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Set Improvement = improvement_list(CStr(classImprovement))
End Property


Property Get uk() As uk_class
' ----------------------------------------------------------------------------
' ����������� ��������
' Last update: 04.04.2018
' ----------------------------------------------------------------------------
    Set uk = uk_list(CStr(classUK))
End Property


Property Get street() As address_street_class
' ----------------------------------------------------------------------------
' �����
' Last update: 18.04.2016
' ----------------------------------------------------------------------------
    Set street = New address_street_class
    street.initial classStreet
End Property


Property Get outReport() As Boolean
' ----------------------------------------------------------------------------
' ����� �� �������� ���� ��� � �����
' Last update: 18.04.2016
' ----------------------------------------------------------------------------
    outReport = classOutReport
End Property


Property Get BldnType() As Long
' ----------------------------------------------------------------------------
' ��� ����
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    BldnType = classBldnType
End Property


Property Get DisRepair() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    DisRepair = classDisRepair
End Property


Property Get EnergoClass() As id_name_class
' ----------------------------------------------------------------------------
' ����� �������������������
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    Set EnergoClass = energo_classes(CStr(classEnergoClass))
End Property


Property Get Address() As String
' ----------------------------------------------------------------------------
' About: ����� �������
' Last update: 31.03.2016
' ----------------------------------------------------------------------------
    ' ��������� �����
    Address = Me.street.Village.Name + " "
    ' ����� (���� ���� ��������)
    Address = Address + IIf(Me.street.Name = NOTSTRING, "", Me.street.Name)
    ' ����� ����
    Address = Address + " �." + classBldnNo
End Property


Property Get AddressWOTown() As String
' ----------------------------------------------------------------------------
' About: ����� ������� ��� ������ ������
' Last update: 31.05.2021
' ----------------------------------------------------------------------------
    ' ��������� �����
    If Not InStr(1, Me.street.Village.Name, "�����", vbTextCompare) > 0 Then
        AddressWOTown = Me.street.Village.Name + " "
    End If
    ' ����� (���� ���� ��������)
    AddressWOTown = AddressWOTown + IIf(Me.street.Name = NOTSTRING, "", Me.street.Name)
    ' ����� ����
    AddressWOTown = AddressWOTown + " �." + classBldnNo
End Property


Property Get ContractNo() As String
' ----------------------------------------------------------------------------
' ����� ��������
' Last update: 25.05.2018
' ----------------------------------------------------------------------------
    ContractNo = classContractNo
End Property


Property Get ContractDate() As Date
' ----------------------------------------------------------------------------
' ���� ��������
' Last update: 25.05.2018
' ----------------------------------------------------------------------------
    ContractDate = classContractDate
End Property


Property Get SubaccountSum() As Currency
' ----------------------------------------------------------------------------
' ������� �� ��������
' Last update: 09.04.2019
' ----------------------------------------------------------------------------
    SubaccountSum = classSubAccount
End Property


Property Get SubaccountDate() As Date
' ----------------------------------------------------------------------------
' ����, �� ������� ������� �� ��������
' Last update: 14.06.2019
' ----------------------------------------------------------------------------
    If classSubAccountDateID = NOTVALUE Or classSubAccountDateID = 0 Then
        SubaccountDate = NOTDATE
    Else
        SubaccountDate = DateAdd("d", 1, terms( _
                                    CStr(classSubAccountDateID)).classEndDate)
    End If
End Property


Property Get CurrentSubaccountSum() As Currency
' ----------------------------------------------------------------------------
' ������� �� �������� � ������ �����
' Last update: 19.02.2020
' ----------------------------------------------------------------------------
    CurrentSubaccountSum = classCurrentSubAccount
End Property


Property Get SubaccountPercent() As Integer
' ----------------------------------------------------------------------------
' ������� ������������ �� ���������
' 20.09.2021
' ----------------------------------------------------------------------------
    SubaccountPercent = WorksheetFunction.Max( _
            WorksheetFunction.Min(classSubAccountPercent, 100), 0)
End Property


Property Get SubaccountPlanSum() As Currency
' ----------------------------------------------------------------------------
' �������� ����������� �� ������� � �����
' Last update: 17.06.2019
' ----------------------------------------------------------------------------
    SubaccountPlanSum = classPlanSubAccountSum
End Property


Property Get SubaccountPlanEndSum() As Currency
' ----------------------------------------------------------------------------
' �������� �������� �� ����� ����
' 25.04.2022
' ----------------------------------------------------------------------------
    Dim monthLeft As Integer
    monthLeft = (12 - Month(Me.SubaccountDate) + 1) Mod 12
    SubaccountPlanEndSum = Me.SubaccountSum + Me.SubaccountPlanSum * monthLeft
End Property


Property Get SubaccountPlanEndWithPercentSum() As Currency
' ----------------------------------------------------------------------------
' �������� �������� �� ����� ���� � ������ ������������
' 25.04.2022
' ----------------------------------------------------------------------------
    Dim monthLeft As Integer
    monthLeft = (12 - Month(Me.SubaccountDate) + 1) Mod 12
    SubaccountPlanEndWithPercentSum = Me.SubaccountSum + Round( _
            Me.SubaccountPlanSum * monthLeft * Me.SubaccountPercent / 100, 2)
End Property


Property Get Fias() As String
' ----------------------------------------------------------------------------
' ��� ����
' Last update: 14.05.2019
' ----------------------------------------------------------------------------
    Fias = classFias
End Property


Property Get GisGuid() As String
' ----------------------------------------------------------------------------
' GUID ���� � ���
' Last update: 16.05.2019
' ----------------------------------------------------------------------------
    GisGuid = classGisGuid
End Property


Property Get ManHourCost() As man_hour_cost
' ----------------------------------------------------------------------------
' ��������� ������������
' 30.05.2022
' ----------------------------------------------------------------------------
    Set ManHourCost = classManHourCost
End Property


Property Let BldnNo(newValue As String)
' ----------------------------------------------------------------------------
' ��������� ������ ����
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Call update(newBldnNo:=newValue)
End Property


Property Get Chairman(Optional InTermId As Long = NOTVALUE) As bldn_chairman_sign
' ----------------------------------------------------------------------------
' ���������� � ����������
' 20.10.2022
' ----------------------------------------------------------------------------
    If InTermId = NOTVALUE Then InTermId = terms.LastTerm.Id
        
    Set Chairman = New bldn_chairman_sign
    Chairman.initial classId, InTermId
End Property


Public Sub create(BldnId As Long, streetId As Long, BldnNo As String, _
                                                mcId As Long, Dogovor As Long)
' ----------------------------------------------------------------------------
' ���������� ����
' 26.05.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    If BldnId = NOTVALUE Then BldnId = 0
    
    sqlString = "create_building"
    sqlParams.add "newbldn", BldnId
    sqlParams.add "newstreet", streetId
    sqlParams.add "newbldnno", BldnNo
    sqlParams.add "newmc", mcId
    sqlParams.add "newcontract", Dogovor
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    classBldnNo = BldnNo
    classStreet = streetId
    classBldnNo = BldnNo
    classUK = mcId
    classManagement = Dogovor
    classId = rst!newId
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        If errorNotUnique(Err.Description) Then
            Err.Raise ERROR_NOT_UNIQUE, "building_class.create", _
                                    "����� ��� ��� ����"
        Else
            Err.Raise Err.Number, "building_class.create", Err.Description
        End If
    End If
End Sub


Public Sub updateServices(newHeating As Long, _
                            newHotWater As Long, _
                            newColdWater As Long, _
                            newGas As Long)
' ----------------------------------------------------------------------------
' ��������� �����
' 26.05.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "change_bldn_services"
    sqlParams.add "InItemId", classId
    sqlParams.add "InHw", newHotWater
    sqlParams.add "InCw", newColdWater
    sqlParams.add "InGas", newGas
    sqlParams.add "InHeating", newHeating
    
    DBConnection.RunQuery sqlString, sqlParams
    
    classColdWater = newColdWater
    classHotWater = newHotWater
    classGas = newGas
    classHeating = newHeating
        
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_class.update_services", Err.Description
    End If
End Sub


Public Sub updateCommon(newImprovement As Long, _
                        newBldnType As Long, _
                        newSiteNo As String, _
                        newCadastral As String, _
                        newDisRepair As Variant, _
                        newEnergoClass As Long, _
                        newFias As String, _
                        newGisGuid As String)
' ----------------------------------------------------------------------------
' ��������� ����� ����������
' 26.05.2022
' ----------------------------------------------------------------------------
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    ' ��� ���� 36 ������ ���� ������
    If Len(newFias) <> 36 And Len(newFias) <> 0 And _
                        StrComp(newFias, NOTSTRING, vbBinaryCompare) <> 0 Then
        MsgBox "�������� �������� ���� ����", vbExclamation, "������"
        Exit Sub
    End If
    
    sqlString = "change_bldn_common"
    sqlParams.add "itemId", classId
    sqlParams.add "newimp", newImprovement
    sqlParams.add "newtype", newBldnType
    sqlParams.add "newsite", newSiteNo
    sqlParams.add "newcadastral", newCadastral
    sqlParams.add "newdisrepair", newDisRepair
    sqlParams.add "newenergoclass", IIf(newEnergoClass = NOTVALUE, _
                                                        Null, newEnergoClass)
    sqlParams.add "newfias", Trim(newFias)
    sqlParams.add "newgisguid", Trim(newGisGuid)
    
    DBConnection.RunQuery sqlString, sqlParams
    
    classImprovement = newImprovement
    classBldnType = newBldnType
    classSiteNo = newSiteNo
    classCadastral = newCadastral
    classDisRepair = newDisRepair
    classEnergoClass = newEnergoClass
    classFias = newFias
    classGisGuid = newGisGuid
        
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_class.update_common", Err.Description
    End If
End Sub


Public Sub updateDogovor(newMC As Long, _
                        newContractor As Long, _
                        newDogovor As Long, _
                        newOutReport As Variant, _
                        newManHourMode As Long)
' ----------------------------------------------------------------------------
' ��������� ���������� � ��������
' 30.05.2022
' ----------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    sqlString = "change_bldn_dogovor"
    sqlParams.add "itemId", classId
    sqlParams.add "newmc", newMC
    sqlParams.add "newcontractor", newContractor
    sqlParams.add "newdogovor", newDogovor
    sqlParams.add "newout", newOutReport
    sqlParams.add "NewManHourMode", newManHourMode
    
    DBConnection.RunQuery sqlString, sqlParams
    
    classUK = newMC
    classContractor = newContractor
    classManagement = newDogovor
    classOutReport = newOutReport
    
    Call getManHourCost
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_class.update_dogovor", Err.Description
    End If
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' �������� ����
' 26.05.2022
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "bldn_class.delete", _
                                                            "������ �� �����"
    End If
    
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "delete_building"
    sqlParams.add "InItemId", classId
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "bldn_class.delete", _
                                                    "���������� ������� ���"
        Else
            Err.Raise Err.Number, "bldn_class.delete", Err.Description
        End If
    End If
End Sub


Private Sub getSubAccount()
' ----------------------------------------------------------------------------
' ������� �� ��������
' Last update: 19.02.2020
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    
    Dim sqlParams As Dictionary
    
    On Error GoTo errHandler
    
    Set sqlParams = New Dictionary
    sqlParams.add "InBldnId", classId
   
    Set rst = DBConnection.GetQueryRecordset("get_bldn_subaccount", sqlParams)
    
    classSubAccount = dblValue(rst!OutSaSum)
    classSubAccountDateID = longValue(rst!OutSaDate)
    classCurrentSubAccount = dblValue(rst!OutSaCurSum)
    
    Set sqlParams = New Dictionary
    sqlParams.add "InBldnId", classId
    Set rst = DBConnection.GetQueryRecordset("get_bldn_plan_subaccount", sqlParams)
    classPlanSubAccountSum = dblValue(rst!plan_sum)
    
    Set sqlParams = New Dictionary
    sqlParams.add "InBldnId", classId
    sqlParams.add "InDate", NOTDATE
    Set rst = DBConnection.GetQueryRecordset("bldn_subaccount_percent", sqlParams)
    classSubAccountPercent = CInt(dblValue(rst!OutPercent) * 100)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "building_class.getSubaccount", Err.Description
    End If
    
End Sub

Private Sub getManHourCost()
' ----------------------------------------------------------------------------
' ���������� ��������� ������������
' 30.05.2022
' ----------------------------------------------------------------------------
    
    Set classManHourCost = New man_hour_cost
    classManHourCost.initialBldnCurrent classId

End Sub

Sub flushValues()
' ----------------------------------------------------------------------------
' ��������� ��������
' Last update: 18.04.2016
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
