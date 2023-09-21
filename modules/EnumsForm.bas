Attribute VB_Name = "EnumsForm"
Option Explicit
' ----------------------------------------------------------------------------
' About: номера столбцов для форм
' ----------------------------------------------------------------------------

Public Enum ListFormType
' ----------------------------------------------------------------------------
' что показывать в ListForm
' 19.10.2022
' ----------------------------------------------------------------------------
    lftCounterModels = 1
    lftExpenseItems
    lftContractors
    lftWorkMaterialTypes
    lftFinanceSources
    lftMunicipalDistrict
    lftHCounterPartTypes
    lftTmpCounters
    lftBldnSubaccounts
    lftFlatHistory
    lftCommonPropertyGroup
    lftCommonPropertyElement
    lftCommonPropertyParameter
    lftExpenseGroups
    lftManHourCostModes
    lftManHourCost
    lftFlatAccrueds
    lftChairmansSign
End Enum


Public Enum FormMCEnum
' ----------------------------------------------------------------------------
' столбцы формы УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    fmcID = 0
    fmcName
    fmcReportName
    fmcDirector
    fmcChiefEngineer
    fmcNotManage
    fmcMax
End Enum


Public Enum FormMDEnum
' ----------------------------------------------------------------------------
' столбцы формы муниципальных образований
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    fmdId = 0
    fmdName
    fmdHead
    fmdHeadPosition
    fmdMax = fmdHeadPosition
End Enum


Public Enum FormVillageEnum
' ----------------------------------------------------------------------------
' столбцы формы населённых пунктов
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    fvId = 0
    fvName
    fvMD
    fvSite
    fvMax
End Enum


Public Enum FormStreetEnum
' ----------------------------------------------------------------------------
' столбцы формы улиц
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    fsId = 0
    fsName
    fsVillage
    fsSite
    fsMax
End Enum


Public Enum FormGWTEnum
' ----------------------------------------------------------------------------
' столбцы формы видов ремонтов
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    fgwtId = 0
    fgwtName
    fgwtNote
    fgwtMax
End Enum


Public Enum FormImprovementEnum
' ----------------------------------------------------------------------------
' столбцы формы видов благоустройства
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    fiId = 0
    fiName
    fiShortName
    fiMax
End Enum


Public Enum FormWorkKindEnum
' ----------------------------------------------------------------------------
' столбцы формы видов работ
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    fwId = 0
    fwName
    fwWorkType
    fwMax
End Enum


Public Enum FormEmployeeEnum
' ----------------------------------------------------------------------------
' столбцы формы работников
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    feId = 0
    feLastName
    feFirstName
    feSecondName
    fePosition
    feSignReport
    feMax
End Enum


Public Enum FormPlanWorkEnum
' ----------------------------------------------------------------------------
' столбцы формы планируемых работ
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    fpwId = 0
    fpwGWT
    fpwWK
    fpwStatus
    fpwDate
    fpwSum
    fpwSmetaSum
    fpwNote
    fpwContractor
    fpwMC
    fpwEmployee
    fpwBeginDate
    fpwEndDate
    fpwPrivateNote
    fpwWorkRef
    fpwMax
End Enum


Public Enum FormNameNoteEnum
' ----------------------------------------------------------------------------
' столбцы формы с названием и описанием/сокращенным названием
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    fnnId = 0
    fnnName
    fnnNote
    fnnMax
End Enum


Public Enum FormWorkListEnum
' ----------------------------------------------------------------------------
' поля формы работ по дому
' Last update: 19.07.2019
' ----------------------------------------------------------------------------
    fwlId = 0
    fwlDate
    fwlWT
    fwlWK
    fwlContractor
    fwlDogovor
    fwlSum
    fwlVolume
    fwlSI
    fwlNote
    fwlFSource
    fwlPrintFlag
    fwlMax
End Enum


Public Enum FormOldWorksEnum
' ----------------------------------------------------------------------------
' поля формы старых работ
' Last update: 05.05.2018
' ----------------------------------------------------------------------------
    fowId = 0
    fowName
    fowYear
    fowVolume
    fowSum
    fowNote
    fowOBF
    fowOBN
    fowMax
End Enum


Public Enum FormBldnLastExpenses
' ----------------------------------------------------------------------------
' поля списка структуры
' Last update: 10.04.2019
' ----------------------------------------------------------------------------
    bfleId = 0
    fbleName
    fblePrice
    fblePlanSum
    fbleFactSum
    fbleBldnName
    fbleDate
    fbleMax
End Enum


Public Enum FormBldnServices
' ----------------------------------------------------------------------------
' поля списка услуг в доме
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    bsServiceId = 0
    bsServiceName
    bsModeId
    bsModeName
    bsInputsCount
    bsPossibleCounter
    bsNote
    bsMax
End Enum


Public Enum UserFormModes
' -----------------------------------------------------------------------------
' режим формы пользователей
' Last update: 13.09.2018
' -----------------------------------------------------------------------------
    ufmAdd = 0
    ufmChangeName = 1
    ufmChangePassword = 2
End Enum


Public Enum FormUserList
' ----------------------------------------------------------------------------
' поля списка пользователей
' Last update: 17.09.2018
' ----------------------------------------------------------------------------
    fulId = 0
    fulLogin
    fulFIO
    fulIsActive
    fulMax
End Enum


Public Enum FormCounterModels
' ----------------------------------------------------------------------------
' поля списка моделей приборов учёта
' Last update: 07.05.2019
' ----------------------------------------------------------------------------
    cmfId = 0
    cmfName
    cmfHasDTI
    cmfCI
    cmfMax
End Enum


Public Enum FormWorkMaterialType
' ----------------------------------------------------------------------------
' поля списка материалов
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    fwmtMaterialId = 0
    fwmtMaterialName
    fwmtIsTransport
    fwmtMax
End Enum


Public Enum FormWorkMaterialsEnum
' ----------------------------------------------------------------------------
' поля материалов работы
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    fwmMaterialId = 0
    fwmMaterialName
    fwmMaterialNote
    fwmMaterialCost
    fwmMaterialCount
    fwmMaterialSi
    fwmMaterialSum
    fwmIsTransport
    fwmMax
End Enum


Public Enum FormFSourceEnum
' -----------------------------------------------------------------------------
' поля формы источников финансирования
' Last update: 21.10.2019
' -----------------------------------------------------------------------------
    ffsId = 0
    ffsName
    ffsFromSubaccount
    ffsNote
    ffsMax = ffsNote
End Enum


Public Enum FormHCounterPartType
' -----------------------------------------------------------------------------
' поля формы частей ОДПУ теплоснабжения
' Last update: 02.07.2020
' -----------------------------------------------------------------------------
    fhcptId = 0
    fhcptName
    fhcptMax = fhcptName
End Enum


Public Enum FormTmpCounters
' -----------------------------------------------------------------------------
' поля формы актов допуска
' Last update: 04.09.2020
' -----------------------------------------------------------------------------
    ftcId = 0
    ftcBldnId
    ftcBldnAddress
    ftcName
    ftcActDate
    ftcMax = ftcActDate
End Enum


Public Enum FormFlats
' -----------------------------------------------------------------------------
' поля формы квартир
' 17.08.2021
' -----------------------------------------------------------------------------
    ffFirst
    ffId = ffFirst
    ffFlatNo
    ffTerm
    ffResidental
    ffUninhabitable
    ffRooms
    ffSquare
    ffNote
    ffMax
End Enum


Public Enum FormFullFlats
' -----------------------------------------------------------------------------
' поля формы квартир
' 01.09.2021
' -----------------------------------------------------------------------------
    fffFirst
    fffId = fffFirst
    fffFlatNo
    fffTerm
    fffResidental
    fffUninhabitable
    fffRooms
    fffSquare
    fffNote
    fffCadastralNo
    fffSaldo
    fffMaxShortInfo
    fffShare = fffMaxShortInfo
    fffIsLegalEntity
    fffIsPrivatized
    fffOwnerId
    fffName
    fffDocument
    fffPhone
    fffHasPdConsent
    fffChairman
    fffSenat
    fffSekretar
    fffMax
End Enum


Public Enum FormBldnCommonPropertiesColumns
' ----------------------------------------------------------------------------
' столбцы формы элементов общего имущества в доме
' 12.04.2022
' ----------------------------------------------------------------------------
    fbcpFirst = 0
    fbcpRank = fbcpFirst
    fbcpGroupId
    fbcpElementId
    fbcpParameterId
    fbcpName
    fbcpState
    fbcpIsUsing
    fbcpMax
End Enum

Public Enum FormChairmanSignColumns
' ----------------------------------------------------------------------------
' столбцы формы
' 20.10.2022
' ----------------------------------------------------------------------------
    fcsFirst = 0
    fcsBeginTerm = fcsFirst
    fcsTermId
    fcsBldnId
    fcsOwnerName
    fcsHasSign
    fcsMax
End Enum
