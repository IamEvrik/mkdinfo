VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectionForm 
   Caption         =   "Выберите справочник"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6225
   OleObjectBlob   =   "SelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' выбор справочника
' ----------------------------------------------------------------------------

Private Enum dictionary_type_enum
' ----------------------------------------------------------------------------
' номера справочников
' 03.06.2022
' ----------------------------------------------------------------------------
    dict_work_kind = 0
    dict_md
    dict_village
    dict_street
    dict_gwt
    dict_work_type
    dict_mc
    dict_contractor
    dict_improvement
    dict_wallmaterial
    dict_dogovor
    dict_fsource
    dict_expense_groups
    dict_expenseitems
    dict_bldnexpensename
    dict_services
    dict_counter_models
    dict_manhour_cost_modes
    dict_manhour_cost
    dict_material_types
    dict_tmp_counters
    dict_common_property_group
    dict_common_property_element
    dict_common_property_parameter
End Enum


Private Enum report_type_enum
' ----------------------------------------------------------------------------
' номера отчетов
' 07.02.2023
' ----------------------------------------------------------------------------
    report_TechInfo = 0
    Report_3
    report_Contractors
    report_Works
    report_MWorksMaterials  ' отчет по содержанию с материалами
    report_ContMaters       ' отчет по материалам
    report_Passport
    Report_1                ' отчет по текущему ремонту
    report_AllWorks
    report_SubAccount
    Report_4
    Report_7
    Report_9            ' отчет по работам с учётом источника финансирования
    Report_r_plan_year
    Report_10           ' отчет для планирования с субсчетами
    report_11           ' состояние субсчетов
    report_13           ' оплата ушедших домов
    report_101          ' проверка начислений
    report_101a         ' проверка начислений за период
    report_102          ' отчет по разовым
    report_12           ' процент собираемости по дому
    report_14           ' собираемость за период по домам
    report_201          ' 22 зима
    report_work_comp    ' месячный акт выполненных работ
End Enum


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы - заполнение списка
' 07.02.2023
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & ". Сервер " & AppConfig.DBServer
    With Me.ListBoxList
        If Me.Tag = "dictionary" Then
            .AddItem "Названия работ"
            .AddItem "Муниципальные образования"
            .AddItem "Населённые пункты"
            .AddItem "Улицы"
            .AddItem "Виды ремонтов"
            .AddItem "Типы работ"
            .AddItem "Управляющие компании"
            .AddItem "Подрядные организации"
            .AddItem "Степени благоустройства"
            .AddItem "Материалы стен"
            .AddItem "Договоры"
            .AddItem "Источники финансирования"
            .AddItem "Группы структуры платы"
            .AddItem "Статьи расходов"
            .AddItem "Названия статей расходов в домах"
            .AddItem "Услуги"
            .AddItem "Модели приборов учёта"
            .AddItem "Режимы стоимости человекочаса"
            .AddItem "Стоимость человекочаса"
            .AddItem "Материалы работ"
            .AddItem "Акты допуска ОДПУ"
            .AddItem "Группы элементов общего имущества"
            .AddItem "Виды элементов общего имущества"
            .AddItem "Параметры элементов общего имущества"
        ElseIf Me.Tag = "report" Then
            .AddItem "Техническая информация МКД"
            .AddItem "3. Информация о земельных участках"
            .AddItem "Отчёт по подрядчикам"
            .AddItem "Отчёт по выполненным работам"
            .AddItem "Отчет по содержанию с материалами"
            .AddItem "Отчет по материалам по подрядчикам"
            .AddItem "Паспорта готовности"
            .AddItem "1. Отчёт по текущему ремонту"
            .AddItem "Все работы"
            .AddItem "Отчёт для субсчетов"
            .AddItem "4. План работ"
            .AddItem "7. Список видов работ"
            .AddItem "9. Отчет по работам в разрезе источников финансирования"
            .AddItem "Годовое планирование работ"
            .AddItem "10. Планируемые субсчета"
            .AddItem "11. Текущее состояние субсчетов"
            .AddItem "13. Оплата ушедших домов"
            .AddItem "101. Проверка начислений"
            .AddItem "101a. Проверка начислений за период"
            .AddItem "102. Отчет по разовым начислениям"
            .AddItem "12. Процент собираемости"
            .AddItem "14. Собираемость по домам за период"
            .AddItem "201. ЖКХ зима"
            .AddItem "Месячный акт выполненных работ"
        End If
    End With
End Sub


Private Sub ButtonRun_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки "запустить"
' 07.02.2023
' ----------------------------------------------------------------------------
    If Me.Tag = "dictionary" Then
        Select Case Me.ListBoxList.ListIndex
            Case -1:
            Case dictionary_type_enum.dict_md: Unload Me: Call RunListForm(lftMunicipalDistrict)
            Case dictionary_type_enum.dict_contractor: Unload Me: Call RunListForm(lftContractors)
            Case dictionary_type_enum.dict_gwt: Unload Me: Call RunGlobalWorkTypeForm
            Case dictionary_type_enum.dict_improvement: Unload Me: Call RunImprovementForm
            Case dictionary_type_enum.dict_mc: Unload Me: Call RunUKForm
            Case dictionary_type_enum.dict_street: Unload Me: Call RunStreetForm
            Case dictionary_type_enum.dict_village: Unload Me: Call RunVillageForm
            Case dictionary_type_enum.dict_work_kind: Unload Me: Call RunWorkKindForm
            Case dictionary_type_enum.dict_work_type: Unload Me: Call RunWorkTypeForm
            Case dictionary_type_enum.dict_wallmaterial: Unload Me: Call RunWMForm
            Case dictionary_type_enum.dict_dogovor: Unload Me: Call RunDogovorForm
            Case dictionary_type_enum.dict_fsource: Unload Me: Call RunListForm(lftFinanceSources)
            Case dictionary_type_enum.dict_expense_groups: Unload Me: Call RunListForm(lftExpenseGroups)
            Case dictionary_type_enum.dict_expenseitems: Unload Me: Call RunListForm(lftExpenseItems)
            Case dictionary_type_enum.dict_bldnexpensename: Unload Me: Call RunBldnExpenseNameForm
            Case dictionary_type_enum.dict_services: Unload Me: Call RunServiceForm
            Case dictionary_type_enum.dict_counter_models: Unload Me: Call RunListForm(lftCounterModels)
            Case dictionary_type_enum.dict_material_types: Unload Me: Call RunListForm(lftWorkMaterialTypes)
            Case dictionary_type_enum.dict_tmp_counters: Unload Me: Call RunListForm(lftTmpCounters)
            Case dictionary_type_enum.dict_common_property_group: Unload Me: Call RunListForm(lftCommonPropertyGroup)
            Case dictionary_type_enum.dict_common_property_element: Unload Me: Call RunListForm(lftCommonPropertyElement)
            Case dictionary_type_enum.dict_common_property_parameter: Unload Me: Call RunListForm(lftCommonPropertyParameter)
            Case dictionary_type_enum.dict_manhour_cost_modes: Unload Me: Call RunListForm(lftManHourCostModes)
            Case dictionary_type_enum.dict_manhour_cost: Unload Me: Call RunListForm(lftManHourCost)
        End Select
    ElseIf Me.Tag = "report" Then
        Select Case Me.ListBoxList.ListIndex
            Case -1:
            Case report_type_enum.report_Contractors:
                Unload Me
                Call RunReport1Form("money")
            Case report_type_enum.report_TechInfo: Unload Me: Call RunReport1Form("Report_2")
            Case report_type_enum.Report_3: Unload Me: Call RunReport1Form("Report_3")
            Case report_type_enum.report_Works: Unload Me: Call RunReport1Form("works")
            Case report_type_enum.report_Passport: Unload Me: Call RunBldnSelectionForm(blrPassport)
            Case report_type_enum.Report_1: Unload Me: Call RunReport1Form("Report_1")
            Case report_type_enum.report_AllWorks: Unload Me: Call RunReport1Form("allworks")
            Case report_type_enum.report_SubAccount: Unload Me: Call RunReport1Form("subaccount")
            Case report_type_enum.Report_4: Unload Me: Call RunReport1Form("Report_4")
            Case report_type_enum.Report_7: Unload Me: Call RunReport1Form("Report_7")
            Case report_type_enum.Report_9: Unload Me: Call RunReport1Form("Report_9")
            Case report_type_enum.Report_r_plan_year: Unload Me: Call ReportYearPlan
            Case report_type_enum.Report_10: Unload Me: Call RunReport1Form("Report_10")
            Case report_type_enum.report_MWorksMaterials: Unload Me: Call RunReport1Form("Report_mworkmaterials")
            Case report_type_enum.report_ContMaters: Unload Me: Call RunReport1Form("Report_contmaterials")
            Case report_type_enum.report_101: Unload Me: Call RunReport1Form("Report_101")
            Case report_type_enum.report_101a: Unload Me: Call RunReport1Form("Report_101a")
            Case report_type_enum.report_102: Unload Me: Call RunReport1Form("Report_102")
            Case report_type_enum.report_11: Unload Me: Call RunReport1Form("Report_11")
            Case report_type_enum.report_12: Unload Me: Call RunReport1Form("Report_12")
            Case report_type_enum.report_14: Unload Me: Call RunReport1Form("Report_14")
            Case report_type_enum.report_201: Unload Me: Call RunReport1Form("Report_201")
            Case report_type_enum.report_13: Unload Me: Call RunReport1Form("Report_13")
            Case report_type_enum.report_work_comp: Unload Me: Call RunBldnSelectionForm(blrWorkCompletition)
        End Select
    End If
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки "отмена"
' Last update: 07.06.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ListBoxList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' ----------------------------------------------------------------------------
' двойной щелчок по пункту меню эмулирует нажатие на кнопку
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Call ButtonRun_Click
End Sub

