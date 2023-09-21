Attribute VB_Name = "Enums"
Option Explicit


Public Enum MainSheetEnum
' ----------------------------------------------------------------------------
' ������� ����� �� ������� �����
' Last update: 11.07.2017
' ----------------------------------------------------------------------------
    msCode = 1
    msUK
    msMD
    msVillage
    msStreet
    msBldnNo
    msContractor
    msDogovor
    msOutReport
End Enum

Public Enum ComboColumns
' ----------------------------------------------------------------------------
' ������� ComboBox
' Last update: 06.12.2016
' ----------------------------------------------------------------------------
    ccId = 1                ' ���
    ccname = 0              ' ��������
End Enum


Public Enum ReloadComboMethods
' ----------------------------------------------------------------------------
' ����� ComboBox ���� ListBox ��� ����������
' 30.05.2022
' ----------------------------------------------------------------------------
    rcmMd
    rcmVillage
    rcmStreet
    rcmMC
    rcmContractor
    rcmGWT
    rcmWorkType
    rcmWorkKind
    rcmTerm
    rcmTermDESC
    rcmListBldnNoId
    rcmListBldnAddressId
    rcmListBldnAddressIdByMD
    rcmListManagedBldnAddressIdByMD
    rcmListBldnAddressIdByVillage
    rcmListBldnAddressIdByStreet
    rcmImprovement
    rcmDogovor
    rcmWallMaterial
    rcmStreetTypes
    rcmMainContractor
    rcmUsingMainContractor
    rcmPlanStatuses
    rcmPlanStatusesNewWork
    rcmVillageTypes
    rcmGas
    rcmHeating
    rcmHotWater
    rcmColdWater
    rcmEmployees
    rcmFSources
    rcmYesNo
    rcmServices
    rcmExpenseGroups
    rcmExpenseItems
    rcmBldnExpenseName
    rcmBldnExpenseTerms
    rcmServiceModes
    rcmUserRoles
    rcmUsers
    rcmUserHasRoles
    rcmUserHasNoRoles
    rcmAccessTypes
    rcmBldnTypes
    rcmEnergoClasses
    rcmRoleHasAccess
    rcmRoleHasNoAccess
    rcmPlanTerms
    rcmWorkMaterialTypes
    rcmRkcServices
    rcmUkServices
    rcmServiceTypes
    rcmFlatTerms
    rcmAddedTypes
    rcmCommonPropertyGroup
    rcmCommonPropertyElement
    rcmSubaccountTerms
    rcmManHourModes
End Enum


Public Enum XmlHeaderTypeEnum
' ----------------------------------------------------------------------------
' ���� ����� ����������� xml ������
' 12.08.2021
' ----------------------------------------------------------------------------
    xhtFlats
End Enum


Public Enum IdNameFormType
' ----------------------------------------------------------------------------
' ��� �������� ��� ����� IdNameForm
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    edfPlanStatus
    edfWallMaterial
    edfWorkType
End Enum


Public Enum BldnListReportType
' ----------------------------------------------------------------------------
' ���� �������, ��������� �� ������ �����
' 27.09.2022
' ----------------------------------------------------------------------------
    blrPassport
    blrWorkCompletition
End Enum


Public Enum PositionStatusEnum
' ----------------------------------------------------------------------------
' ���������
' Last update: 17.04.2018
' ----------------------------------------------------------------------------
    psDirector = 1
    psChiefEngineer = 2
    psOther = 3
End Enum


Public Enum ConstValuesEnum
' ----------------------------------------------------------------------------
' �������� �� ����� Settings
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    cveUserId = 1
End Enum


Public Enum ImportSubAccounts
' ----------------------------------------------------------------------------
' ������� ����� �������� ���������
' Last update: 11.07.2019
' ----------------------------------------------------------------------------
    isMC = 1
    isMD
    isVillage
    isBldnId
    isStreet
    isBldnNo
    isSquare
    isAccrued
    isPaid
    isMonth
End Enum


Public Enum AccruedTypes
' ----------------------------------------------------------------------------
' ��� ���������
' Last update: 23.03.2021
' ----------------------------------------------------------------------------
    acRKC = 1
    acBuh = 2
    acMOKR
    acMOUpper
End Enum


Public Enum YurLicoSheetColumns
' ----------------------------------------------------------------------------
' ������� ���������� �������
' Last update: 25.01.2022
' ----------------------------------------------------------------------------
    ylscName = 1
    ylscAddress
    ylscFlatNo
    ylscBldnId
    ylscOccId
    ylscSodAccrued
    ylscSodPaid
    ylscElectroAccrued
    ylscElectroPaid
    ylscColdWaterAccrued
    ylscColdWaterPaid
    ylscHotWaterAccrued
    ylscHotWaterPaid
End Enum


Public Enum MOKarRemontSheetColumns
' ----------------------------------------------------------------------------
' ������� ���������� ������ �� ��������� ��
' 13.04.2022
' ----------------------------------------------------------------------------
    mkrUK = 1
    mkrMO
    mkrAddress
    mkrBldnId
    mkrFlatNo
    mkrOccId
    mkrAccrued
    mkrPaid
End Enum


Public Enum MoUpperSheetColumns
' ----------------------------------------------------------------------------
' ������� ���������� ���������� ��
' Last update: 25.03.2021
' ----------------------------------------------------------------------------
    muFirst = 1
    muPP = muFirst
    muAddress
    muBldnId
    muMuniSquare
    muOwnerSod
    muMuniSod
    muSodUpper
    muSodAccured
    muOwnerColdWater
    muMuniColdWater
    muColdWaterUpper
    muColdWaterAccrued
    muOwnerHotWater
    muMuniHotWater
    muHotWaterUpper
    muHotWaterAccrued
    muTotalUpper
    muTotalAccrued
    muSodPaid
    muColdWaterPaid
    muHotWaterPaid
    muLast = muHotWaterPaid
End Enum

