Attribute VB_Name = "Enums"
Option Explicit


Public Enum MainSheetEnum
' ----------------------------------------------------------------------------
' столбцы листа со списком домов
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
' столбцы ComboBox
' Last update: 06.12.2016
' ----------------------------------------------------------------------------
    ccId = 1                ' код
    ccname = 0              ' название
End Enum


Public Enum ReloadComboMethods
' ----------------------------------------------------------------------------
' выбор ComboBox либо ListBox для заполнения
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
' виды типов заголовоков xml файлов
' 12.08.2021
' ----------------------------------------------------------------------------
    xhtFlats
End Enum


Public Enum IdNameFormType
' ----------------------------------------------------------------------------
' тип объектов для формы IdNameForm
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    edfPlanStatus
    edfWallMaterial
    edfWorkType
End Enum


Public Enum BldnListReportType
' ----------------------------------------------------------------------------
' виды отчетов, выводимых по списку домов
' 27.09.2022
' ----------------------------------------------------------------------------
    blrPassport
    blrWorkCompletition
End Enum


Public Enum PositionStatusEnum
' ----------------------------------------------------------------------------
' должности
' Last update: 17.04.2018
' ----------------------------------------------------------------------------
    psDirector = 1
    psChiefEngineer = 2
    psOther = 3
End Enum


Public Enum ConstValuesEnum
' ----------------------------------------------------------------------------
' значения на листе Settings
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    cveUserId = 1
End Enum


Public Enum ImportSubAccounts
' ----------------------------------------------------------------------------
' столбцы книги загрузки субсчетов
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
' Кто начисляет
' Last update: 23.03.2021
' ----------------------------------------------------------------------------
    acRKC = 1
    acBuh = 2
    acMOKR
    acMOUpper
End Enum


Public Enum YurLicoSheetColumns
' ----------------------------------------------------------------------------
' Столбцы начислений юрлицам
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
' Столбцы начислений взноса на капремонт МО
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
' Столбцы начислений превышения МО
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

