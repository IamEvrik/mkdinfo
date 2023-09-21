Attribute VB_Name = "EnumReports"
Option Explicit

Public Enum TSREnum
' ----------------------------------------------------------------------------
' номера столбцов отчета на сайт
' Last update: 22.08.2016
' ----------------------------------------------------------------------------
    tsrName = 1
    tsrTotal
    tsrTR
    tsrWorks
    tsr01
    tsr02
    tsr03
    tsr04
    tsr05
    tsr06
    tsr07
    tsr08
    tsr09
    tsr10
    tsr11
    tsr12
    tsrLast = tsr12
End Enum


Public Enum TotalReportEnum
' -----------------------------------------------------------------------------
' номера столбцов листа общего отчета по выполнению
' Last update: 31.03.2016
' -----------------------------------------------------------------------------
    treNPP = 1
    treBldn
    treContractor
    treMC
    treAddress
    treSquare
    treAVR
    treAccruedMonth
    treYearPlan
    treAccrued
    treWorks
    treDifference
    trePercent
End Enum


Public Enum WorkReportEnum
' -----------------------------------------------------------------------------
' номера столбцов листа отчета по работам
' Last update: 15.09.2016
' -----------------------------------------------------------------------------
    wrePP = 1
    wreBldn
    wreContractor
    wreAddress
    wreWorkType
    wreWork
    wreDogovor
    wreVolume
    wreSum
    wreDate
End Enum


Public Enum Report1Enum
' -----------------------------------------------------------------------------
' номера столбцов листа отчета report1
' Last update: 29.10.2019
' -----------------------------------------------------------------------------
    r1ePP = 1
    r1eMC
    r1eAddress
    r1eBldnId
    r1eWT
    r1eWork
    r1eContractor
    r1eDogovor
    r1eSI
    r1eVolume
    r1eSum
    r1eLast = r1eSum
    r1eDirString = 2
    r1eDirFIO = 7
End Enum


Public Enum Report2Enum
' -----------------------------------------------------------------------------
' номера столбцов листа отчета report2 (техническая информация)
' Last update: 26.01.2021
' -----------------------------------------------------------------------------
    r2ePP = 1
    r2eID
    r2eAddress
    r2eContractor
    r2eDogovor
    r2eHeating
    r2eHotWater
    r2eGas
    r2eHasOdpuHW
    r2eHasOdpuHeating
    r2eHasOdpuCommon
    r2eHasOdpuCW
    r2eHasOdpuEE
    r2eHasThermoregulator
    r2eFloorMin
    r2eFloorMax
    r2eWallMaterial
    r2eBuiltYear
    r2eCommissioningYear
    r2eHasDoorPhone
    r2eHasDoorCloser
    r2eTotalSquare
    r2eStairsSquare
    r2eCorridorSquare
    r2eOtherMOPSquare
    r2eMOPSquare
    r2eSquareBanisters
    r2eSquareDoors
    r2eSquareWindowSills
    r2eSquareDoorHandles
    r2eSquareMailBoxes
    r2eSquareRadiatorsMOP
    r2eEntrancesCount
    r2eStairsCount
    r2eVaultsCount
    r2eVaultsSquare
    r2eAtticSquare
    r2eStructuralVolume
    r2eDepreciation
End Enum


Public Enum ReportAllWorks
' -----------------------------------------------------------------------------
' номера столбцов листа работ за все годы
' Last update: 30.01.2017
' -----------------------------------------------------------------------------
    rawPP = 1
    rawBldnId
    rawAddress
    rawWork
    rawYear
    rawVolume
    rawSum
    rawBudget
    rawNote
End Enum


Public Enum SubAccountReportEnum
' -----------------------------------------------------------------------------
' номера столбцов листа отчета для субсчетов
' Last update: 26.01.2018
' -----------------------------------------------------------------------------
    sarBldn = 1
    sarSum
    sarDate
    sarVolume
    sarWorkName
    sarContractor
    sarDogovor
    sarNote
    sarVolumeOnly
    sarSi
End Enum


Public Enum BldnWorkReportEnum
' -----------------------------------------------------------------------------
' номера столбцов отчёта для сайта без сумм
' Last update: 29.01.2018
' -----------------------------------------------------------------------------
    bwrContractor = 1
    bwrGWT
    bwrWK
    bwrVolume
    bwrDate
    bwrLast = bwrDate
End Enum


Public Enum Report3Enum
' -----------------------------------------------------------------------------
' номера столбцов отчёта №3 (земельный участок)
' Last update: 15.05.2018
' -----------------------------------------------------------------------------
    r3ID = 1
    r3Address
    r3Contract
    r3Cadastral
    r3Inventory
    r3Use
    r3Survey
    r3Builup
    r3Undeveloped
    r3Hard
    r3DriveWays
    r3SideWalks
    r3OtherHard
    r3SAF
    r3Fences
    r3Benches
End Enum


Public Enum Report4Enum
' -----------------------------------------------------------------------------
' номера столбцов отчёта №4 (план работ)
' Last update: 28.01.2019
' -----------------------------------------------------------------------------
    r4Mc = 1
    r4Addredd
    r4GWT
    r4WorkKind
    r4Contractor
    r4PlanDate
    r4PlanSum
    r4SmetaSum
    r4PlanBDate
    r4PlanEDate
    r4Status
    r4Employee
    r4Last = r4Employee
End Enum


Public Enum Report7Enum
' -----------------------------------------------------------------------------
' номера столбцов отчёта №7 (список видов работ)
' Last update: 10.05.2018
' -----------------------------------------------------------------------------
    r7WkId = 1
    r7WtName
    r7WkName
    r7Last = r7WkName
End Enum


Public Enum Report9Enum
' -----------------------------------------------------------------------------
' номера столбцов листа отчета report_9
' Last update: 09.07.2020
' -----------------------------------------------------------------------------
    r9ePP = 1
    r9eMC
    r9eAddress
    r9eBldnId
    r9eWT
    r9eWork
    r9eContractor
    r9eDogovor
    r9eFSource
    r9eVolume
    r9eSI
    r9eSum
    r9eLast = r9eSum
End Enum


Public Enum ReportBldnWorksEnum
' -----------------------------------------------------------------------------
' номера столбцов листа отчета работ по дому
' Last update: 28.05.2019
' -----------------------------------------------------------------------------
    rbwDate = 1
    rbwContractor
    rbwWorkKind
    rbwVolume
    rbwDogovor
    rbwSum
    rbwFSource
    rbwLast = rbwFSource
End Enum


Public Enum ReportYearPlanCol
' -----------------------------------------------------------------------------
' номера столбцов отчета для планирования работ
' Last update: 29.07.2019
' -----------------------------------------------------------------------------
    repBldnId = 1
    repAddress
    repWorkName
    repContractorName
    repMonthName
    repWorkSum
    repWorkStatus
    repCurrentSacc
    repPlanEndSacc
    repPlanEndWork
    repM1
    repM2
    repM3
    repM4
    repM5
    repM6
    repM7
    repM8
    repM9
    repM10
    repM11
    repM12
    repLast = repM12
End Enum


Public Enum ReportSubAccountsPlanCol
' -----------------------------------------------------------------------------
' номера столбцов отчета для еще одного планирования работ
' Last update: 26.08.2020
' -----------------------------------------------------------------------------
    rsapBlndId = 1
    rsapMC
    rsapContractor
    rsapAddress
    rsapSquare
    rsapCurrentMoney
    rsapWorks
    rsapWorkSum
    rsapPlanMonth
    rsapPercent
    rsapPlanPaids
    rsapYearEnd
    rsapTR
    rsapKR
    rsapPlanMonthNextYear
    rsapFactMonthNextYear
    rsapLast = rsapKR
End Enum


Public Enum ReportMWorkMaterialCol
' -----------------------------------------------------------------------------
' номера столбцов отчета по содержанию с материалами
' Last update: 31.10.2019
' -----------------------------------------------------------------------------
    rmwmcBldnId = 1
    rmwmcContractor
    rmwmcAddress
    rmwmcWork
    rmwmcMaterial
    rmwmcTransport
    rmwmcManHours
    rmwmcWorkDate
End Enum


Public Enum ReportContractorMaterialCol
' -----------------------------------------------------------------------------
' номера столбцов отчета по материалам
' Last update: 31.10.2019
' -----------------------------------------------------------------------------
    rcmcContractor = 1
    rcmcMaterialName
    rcmcMaterialSum
    rcmcTransport
End Enum


Public Enum Report101CheckAccrueds
' -----------------------------------------------------------------------------
' номера столбцов отчета проверки начислений
' 12.10.2021
' -----------------------------------------------------------------------------
    rep101First = 1
    rep101BldnId = rep101First
    rep101Address
    rep101Square
    rep101Price
    rep101Accrued
    rep101Added
    rep101AddedCom
    rep101AddedClean
    rep101AddedDolg
    rep101AddedDiff
    rep101Diff
    rep101Paid
    rep101Compens
    rep101Last = rep101Compens
End Enum


Public Enum Report102
' -----------------------------------------------------------------------------
' номера столбцов отчета истории разовых
' 15.09.2021
' -----------------------------------------------------------------------------
    rep102First = 1
    rep102BldnId = rep102First
    rep102Address
    rep102Term
    rep102Sum
    rep102Last = rep102Sum
End Enum


Public Enum Report11
' -----------------------------------------------------------------------------
' номера столбцов отчета по субсчетам
' 08.10.2021
' -----------------------------------------------------------------------------
    rep11First = 1
    rep11BldnId = rep11First
    rep11PP
    rep11Address
    rep11YearBeginSum
    rep11Sum
    rep11Percent
    rep11PlanPercentSum
    rep11PlanSum
    rep11EndPercentSum
    rep11EndSum
    rep11YearAccrued
    rep11YearPaid
    rep11Last = rep11YearPaid
End Enum


Public Enum Report12
' -----------------------------------------------------------------------------
' номера столбцов собираемость за период по дому
' 30.01.2023
' -----------------------------------------------------------------------------
    rep12First = 1
    rep12OccId = rep11First
    rep12Flat
    rep12FIO
    rep12InSaldo
    rep12Accrued
    rep12Added
    rep12Compens
    rep12Paid
    rep12TotalAccrued
    rep12TotalPaid
    rep12OutSaldo
    rep12Percent
    rep12Warning
    rep12Last = rep12Warning
End Enum


Public Enum ReportInspectionColumns
' -----------------------------------------------------------------------------
' номера столбцов акта обследования
' 05.05.2022
' -----------------------------------------------------------------------------
    ricFirst = 1
    ricName = ricFirst
    ricState
    ricLast = ricState
End Enum


Public Enum BldnCPEColumns
' -----------------------------------------------------------------------------
' номера столбцов состава общего имущества
' 25.05.2022
' -----------------------------------------------------------------------------
    bcpFirst = 1
    bcpName = bcpFirst
    bcpParameter
    bcpState
    bcpLast = bcpState
End Enum


Public Enum Report201Column
' -----------------------------------------------------------------------------
' номера столбцов отчета 201 (ЖКХ зима)
' 07.06.2022
' -----------------------------------------------------------------------------
    r201First = 1
    r201BldnId = r201First
    r201MD
    r201Address
    r201Square
    r201WorkSum
    r201Last = r201WorkSum
End Enum


Public Enum ReportWCompl
' -----------------------------------------------------------------------------
' номера столбцов месячного акта выполненных работ
' 22.09.2022
' -----------------------------------------------------------------------------
    rwcFirst = 1
    rwcPP = rwcFirst
    rwcName
    rwcSum
    rwcLast = rwcSum
End Enum


Public Enum Report13Column
' -----------------------------------------------------------------------------
' номера столбцов отчета 13 (оплата по снятым домам)
' 20.12.2022
' -----------------------------------------------------------------------------
    r13First = 1
    r13BldnId = r13First
    r13Address
    r13Term
    r13Service
    r13Sum
    r13Last = r13Sum
End Enum


Public Enum Report14Column
' -----------------------------------------------------------------------------
' номера столбцов отчета 14 (собираемость по домам)
' 13.02.2023
' -----------------------------------------------------------------------------
    r14First = 1
    r14BldnId = r14First
    r14Address
    r14Accrued
    r14Compens
    r14Addeds
    r14DolgAddeds
    r14ClearAddeds
    r14Paid
    r14Percent
    r14FullPercent
    r14Last = r14FullPercent
End Enum
