Attribute VB_Name = "Constants"
Option Explicit

Public Const appName As String = "ukworks"

' константы базы данных
Public Const DB_ADM_NAME = "postgres"
Public Const DB_NAME = "totalmkd"
Public Const DB_DRIVER = "{PostgreSQL UNICODE}"
Public Const DB_ADM_UID = "postgres"
Public Const DB_PORT = "5432"
Public Const DB_UID = "yp_mc"
Public Const DB_PWD = "yppwd"
Public Const FIRST_PLAN_YEAR = 2018

'Public Const DB_ADM_PWD = "rjvgfybz1"
Public Const DB_ADM_PWD = "17063561"

Public Const shtnmTitul As String = "Титульный лист"
Public Const shtnmMain As String = "Список домов"
Public Const shtSettings As String = "Settings"

' -----------------------------------------------------------------------------
' Названия базы данных, таблиц и представлений
' -----------------------------------------------------------------------------
Public Const MDTableName As String = "municipal_districts"
Public Const VillageTypeTableName As String = "village_types"
Public Const VillageTableName As String = "villages"
Public Const StreetTypeTableName As String = "street_types"
Public Const StreetTableName As String = "streets"
Public Const GWTTableName As String = "global_work_types"
Public Const WTTableName As String = "work_types"
Public Const WKTableName As String = "work_kinds"
Public Const ContractorTableName As String = "contractors"
Public Const ImprovementTableName As String = "improvements"
Public Const EmployeeTableName As String = "employees"
Public Const MCTableName As String = "management_companies"
Public Const BuildingTableName As String = "buildings"
Public Const WorkTableName As String = "works"
Public Const BldnLandInfoTableName As String = "buildings_land_info"
Public Const BldnTechInfoTableName As String = "buildings_tech_info"
Public Const AccruedTableName As String = "accrueds"
Public Const TermTableName As String = "terms"
Public Const WallMaterialTableName As String = "wall_materials"
Public Const PlanChangeReasonTableName As String = "plan_change_reasons"
Public Const InspectionTableName As String = "inspections"
Public Const DogovorTableName As String = "dogovors"
Public Const ConstantsTableName As String = "constants"
Public Const PlanWorkTableName As String = "plan_works"
Public Const PlanStatusTableName As String = "plan_work_statuses"
Public Const PositionStatusTableName As String = "position_statuses"
Public Const FSourceTableName As String = "work_financing_sources"
Public Const ExpenseItemsTableName As String = "expense_items"

Public Const BldnWorksViewName As String = "buildings_workslist"

' ----------------------------------------------------------------------------
' Названия файлов sql скриптов
' ----------------------------------------------------------------------------
Public Const CreateTablesQueryFile = "create_tables.sql"

' ----------------------------------------------------------------------------
' Названия файлов шаблонов отчётов
' ----------------------------------------------------------------------------
Public Const CommonPropertiesFile = "pril2.dotx"
Public Const TemplatePlanList = "план-лист.xltx"

' ----------------------------------------------------------------------------
' Названия файлов шаблонов ГИС
' ----------------------------------------------------------------------------
Public Const GIS_PLAN_EXPENSES = "Экспорт перечня работ.xltx"

' Код глобального типа работ по содержанию
Public Const SERVICE_GLOBAL_TYPE As Long = 1

' Пароли на лист и VBA
Public Const shtPass As String = "1qazxsw2"
Public Const VBA_PASS As String = "17063561"

' -----------------------------------------------------------------------------
' Константы для указания, что значение на задано
' -----------------------------------------------------------------------------
Public Const NOTVALUE As Integer = -1001
Public Const NOTSTRING As String = "-------"
Public Const NOTDATE = #1/1/1900#
Public Const ALL_STRING As String = "-----Все-----"
Public Const ALLVALUES As Integer = -1002
Public Const OTHERVALUE As Integer = -1111


' -----------------------------------------------------------------------------
' Константы Word
' -----------------------------------------------------------------------------
Public Const wdFormatOriginalFormatting = 16
Public Const wdReplaceAll = 2
Public Const wdFindContinue = 1
Public Const wdDoNotSaveChanges = 0
Public Const wdGoToBookmark = -1
Public Const wdStory = 6
Public Const wdExtend = 1
Public Const wdExportFormatPDF = 17
Public Const wdExportOptimizeForOnScreen = 1

' -----------------------------------------------------------------------------
' Константы форм
' -----------------------------------------------------------------------------
Public Const FRM_TEXT_ALIGN_LEFT = 1
Public Const FRM_TEXT_ALIGN_CENTER = 2
Public Const FRM_TEXT_ALIGN_RIGHT = 3

Public Const FRM_ALIGNMENT_LEFT = 0
Public Const FRM_ALIGNMENT_RIGHT = 1
