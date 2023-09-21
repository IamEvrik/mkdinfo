VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuildingForm 
   Caption         =   "Информация по дому"
   ClientHeight    =   11250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19905
   OleObjectBlob   =   "BuildingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BuildingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' переменные формы
' ----------------------------------------------------------------------------
Public formBldnId As Long                   ' для передачи начального значения
Private curWork As work_class               ' выбранная работа
Private curPlanWork As plan_work_class      ' выбранная плановая работа
Private formPrefGWT As Long                 ' предпочитаемый вид ремонта
Private enableEvents As Boolean             ' выбирать ли пункт списка
                                            ' при обновлении
Private curItem As New building_class       ' выбранный дом
Private curTechInfo As New bldnTechInfo     ' его техническая информация
Private curLandInfo As New bldnLandInfo     ' его информация о З/У
Private selPage As Long                     ' активная вкладка, при смене дома
                                            ' выбирается та же вкладка
Public workChanged As Boolean               ' флаг необходимости обновления
                                            ' списка работ
Public planWorkChanged As Boolean           ' флаг необходимости обновления
                                            ' плана работ
Private oldWorkChanged As Boolean           ' флаг необходимости обновления
                                            ' старых работ
Private updateExpenses As Boolean           ' флаг необходимость обновления
                                            ' расходов
Private updateFlats As LoadStatesEnum       ' флаг необходимость обновления
                                            ' квартир
Private oldWorks As old_works
Private curOldWork As old_work
Public formIsActive As Boolean

Private Enum InfoPagesEnum
' ----------------------------------------------------------------------------
' номера вкладок
' 26.11.2021
' ----------------------------------------------------------------------------
    ipInfo = 0
    ipFlats
    ipTechInfo
    ipCommonPropertyElements
    ipLand
    ipWorks
    ipPlanWorks
    ipOldWorks
    ipExpense
    ipOffersWorks
End Enum

Private Enum LoadStatesEnum
' ----------------------------------------------------------------------------
' статусы загрузки для обновления данных
' 20.08.2021
' ----------------------------------------------------------------------------
    ls_clean
    ls_data
    ls_full
End Enum


' ----------------------------------------------------------------------------
' константы названий кнопок
' ----------------------------------------------------------------------------
Const CHANGE_CAPTION = "Изменить"
Const SAVE_CAPTION = "Сохранить"


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' инициализация формы
' 30.05.2022
' ----------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandler
    
    Me.Caption = "Информация по дому (" & AppConfig.AppVersion & "). " & _
                                "База данных: " & DBConnection.ServerAddress
    
    ' прячем элемент с информацией о доме
    Me.InfoPages.Visible = False
    ' по умолчанию первая вкладка
    selPage = 0
    
    ' предпочитаемый вид работ
    formPrefGWT = getPrefetchWork
    
    ' устанавливаем флаг, что нужно обновить список работ
    workChanged = True
    planWorkChanged = True
    oldWorkChanged = True
    
    ' заполнение списка МО с установкой флага не выбирать значение
    enableEvents = False
    Call reloadComboBox(rcmMd, Me.ComboBoxMO)
    
    ' заполняем списки, которые в принципе не должны меняться
    Call reloadComboBox(rcmGWT, Me.ComboBoxGWT, defValue:=formPrefGWT)
    Call reloadComboBox(rcmWallMaterial, Me.ComboBoxWallMaterial)
    Call reloadComboBox(rcmMC, Me.ComboBoxMCChange)
    Call reloadComboBox(rcmUsingMainContractor, Me.ComboBoxContractorChange)
    Call reloadComboBox(rcmImprovement, Me.ComboBoxImprovement)
    Call reloadComboBox(rcmDogovor, Me.ComboBoxDogovor)
    Call reloadComboBox(rcmGas, Me.ComboBoxGas)
    Call reloadComboBox(rcmHeating, Me.ComboBoxHeating)
    Call reloadComboBox(rcmHotWater, Me.ComboBoxHotWater)
    Call reloadComboBox(rcmColdWater, Me.ComboBoxColdWater)
    Call reloadComboBox(rcmExpenseItems, Me.ComboBoxExpenseNames, _
                                                            addAllItems:=True)
    Call reloadComboBox(rcmBldnTypes, Me.ComboBoxBldnType)
    Call reloadComboBox(rcmEnergoClasses, Me.ComboBoxEnergoClass)
    Call reloadComboBox(rcmManHourModes, Me.ComboBoxManHourMode)
    
    ' формирование вида ListView с планами работ
    Me.ListViewPlanWork.View = lvwReport
    Me.ListViewPlanWork.Gridlines = True
    Me.ListViewPlanWork.FullRowSelect = True
    Me.ListViewPlanWork.LabelEdit = lvwManual
    With Me.ListViewPlanWork.ColumnHeaders
        .Clear
        For i = 1 To FormPlanWorkEnum.fpwMax
            .add
        Next i
        .Item(FormPlanWorkEnum.fpwContractor + 1).text = "Подрядчик"
        .Item(FormPlanWorkEnum.fpwDate + 1).text = "Дата"
        .Item(FormPlanWorkEnum.fpwEmployee + 1).text = "Ответственный"
        .Item(FormPlanWorkEnum.fpwGWT + 1).text = "Ремонт"
        .Item(FormPlanWorkEnum.fpwId + 1).text = "Код"
        .Item(FormPlanWorkEnum.fpwMC + 1).text = "УК"
        .Item(FormPlanWorkEnum.fpwNote + 1).text = "Примечание"
        .Item(FormPlanWorkEnum.fpwPrivateNote + 1).text = "Комментарий"
        .Item(FormPlanWorkEnum.fpwStatus + 1).text = "Статус"
        .Item(FormPlanWorkEnum.fpwSum + 1).text = "Сумма"
        .Item(FormPlanWorkEnum.fpwWK + 1).text = "Работа"
        .Item(FormPlanWorkEnum.fpwWorkRef + 1).text = "Код работы"
        .Item(FormPlanWorkEnum.fpwBeginDate + 1).text = "Начало работы"
        .Item(FormPlanWorkEnum.fpwEndDate + 1).text = "Конец работы"
        .Item(FormPlanWorkEnum.fpwSmetaSum + 1).text = "Сумма по смете"
    End With
    
    ' формирование вида ListView с работами
    With Me.ListViewList
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Clear
            For i = 1 To FormWorkListEnum.fwlMax
                .add
            Next i
            .Item(FormWorkListEnum.fwlContractor + 1).text = "Подрядчик"
            .Item(FormWorkListEnum.fwlDate + 1).text = "Дата"
            .Item(FormWorkListEnum.fwlDogovor + 1).text = "Договор"
            .Item(FormWorkListEnum.fwlFSource + 1).text = "Ист. финансирования"
            .Item(FormWorkListEnum.fwlId + 1).text = "Код"
            .Item(FormWorkListEnum.fwlNote + 1).text = "Примечание"
            .Item(FormWorkListEnum.fwlPrintFlag + 1).text = "В отчёт"
            .Item(FormWorkListEnum.fwlSI + 1).text = "Ед.изм."
            .Item(FormWorkListEnum.fwlSum + 1).text = "Сумма"
            .Item(FormWorkListEnum.fwlVolume + 1).text = "Объем"
            .Item(FormWorkListEnum.fwlWK + 1).text = "Работа"
            .Item(FormWorkListEnum.fwlWT + 1).text = "Вид работы"
        End With
    End With
    
    ' формирование вида ListView со старыми работами
    With Me.ListViewOldWorks
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Clear
            For i = 1 To FormOldWorksEnum.fowMax
                .add
            Next i
            .Item(FormOldWorksEnum.fowId + 1).text = "код"
            .Item(FormOldWorksEnum.fowName + 1).text = "Работа"
            .Item(FormOldWorksEnum.fowNote + 1).text = "Примечание"
            .Item(FormOldWorksEnum.fowOBF + 1).text = "др.фин."
            .Item(FormOldWorksEnum.fowOBN + 1).text = "финансы"
            .Item(FormOldWorksEnum.fowSum + 1).text = "Сумма"
            .Item(FormOldWorksEnum.fowVolume + 1).text = "Объем"
            .Item(FormOldWorksEnum.fowYear + 1).text = "Год"
        End With
    End With
        
    ' формирование вида ListView со структурой
    With Me.ListViewExpenses
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Clear
            For i = 1 To FormBldnLastExpenses.fbleMax
                .add
            Next i
            .Item(FormBldnLastExpenses.bfleId + 1).text = "Код"
            .Item(FormBldnLastExpenses.fbleBldnName + 1).text = "Название для дома"
            .Item(FormBldnLastExpenses.fbleName + 1).text = "Название"
            .Item(FormBldnLastExpenses.fblePrice + 1).text = "Цена"
            .Item(FormBldnLastExpenses.fblePlanSum + 1).text = "Сумма план"
            .Item(FormBldnLastExpenses.fbleFactSum + 1).text = "Сумма факт"
            .Item(FormBldnLastExpenses.fbleDate + 1).text = "Период"
        End With
    End With
        
    ' формирование вида ListView с услугами
    With Me.ListViewServices
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        With .ColumnHeaders
            .Clear
            For i = 1 To FormBldnServices.bsMax
                .add
            Next i
            .Item(FormBldnServices.bsInputsCount + 1).text = "Количество вводов"
            .Item(FormBldnServices.bsModeId + 1).text = "Код режима"
            .Item(FormBldnServices.bsModeName + 1).text = "Режим"
            .Item(FormBldnServices.bsServiceId + 1).text = "Код услуги"
            .Item(FormBldnServices.bsServiceName + 1).text = "Услуга"
            .Item(FormBldnServices.bsPossibleCounter + 1).text = "Возможность ОДПУ"
            .Item(FormBldnServices.bsNote + 1).text = "Примечание"
        End With
        .Top = 6
        .Left = 6
    End With
    Me.Top = 1
'    Dim nHeight As Double
'    Dim nZoom As Double
'    nHeight = Application.UsableHeight
'    nZoom = nHeight / Me.Height
'    Me.Height = Application.UsableHeight
'    Me.Zoom = nZoom * 100
    
    ' масштабирование по вертикали
    Dim hDiff As Double
    Dim maxhSize As Double
    maxhSize = GetSystemMetrics(1) * 0.75 - 40
    hDiff = maxhSize - Me.Height
    Me.Height = maxhSize
    
    Me.InfoPages.Height = Me.InfoPages.Height + hDiff
    Me.ListViewCPE.Height = Me.ListViewCPE.Height + hDiff
    Me.ListViewExpenses.Height = Me.ListViewExpenses.Height + hDiff
    Me.ListViewFlats.Height = Me.ListViewFlats.Height + hDiff
    Me.ListViewPlanWork.Height = Me.ListViewPlanWork.Height + hDiff
    Me.ListViewList.Height = Me.ListViewList.Height + hDiff
    
    ' масштабирование по горизонтали
    Dim wDiff As Double
    Dim maxwSize As Double
    maxwSize = GetSystemMetrics(0) * 0.75 - 40
    wDiff = maxwSize - Me.Width
    Me.Width = maxwSize
    
    Me.InfoPages.Width = Me.InfoPages.Width + wDiff
    Me.ListViewCPE.Width = Me.ListViewCPE.Width + wDiff
    Me.ListViewExpenses.Width = Me.ListViewExpenses.Width + wDiff
    Me.ListViewFlats.Width = Me.ListViewFlats.Width + wDiff
    Me.ListViewPlanWork.Width = Me.ListViewPlanWork.Width + wDiff
    Me.ListViewList.Width = Me.ListViewList.Width + wDiff
    GoTo cleanHandler
    
    
errHandler:
    MsgBox Err.Description
    End
    
cleanHandler:
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы, заполнение нужной информации
' Last update: 02.05.2018
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    formIsActive = True
    ' перед заполнением списка МО проверяем, выбран ли дом,
    ' если выбран, то устанавливаем флаг выбора адреса
    If formBldnId <> NOTVALUE And formBldnId <> 0 Then
        enableEvents = True
        Call initialCurItem
        Call selectComboBoxValue(Me.ComboBoxMO, _
                                curItem.street.Village.Municipal_district.Id)
    End If
    If selPage <> 0 And selPage <> NOTVALUE Then
        Call loadInfo(selPage)
    End If
    ' убираем флаг выбора адреса и обнуляем код дома, чтобы не лез обновляться
    formBldnId = NOTVALUE
    enableEvents = False
    GoTo cleanHandler
    
errHandler:
    MsgBox Err.Description
    End
    
cleanHandler:
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' уничтожение формы
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Call terminateVars
End Sub


Private Sub ComboBoxMO_Change()
' ----------------------------------------------------------------------------
' обработка выбора муниципального образования
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxMO.ListIndex > -1 Then
        Dim dfValue As Long
        Me.InfoPages.Visible = False
        Me.ComboBoxStreet.Clear
        Me.ComboBoxBldn.Clear
        If enableEvents Then
            dfValue = curItem.street.Village.Id
        Else
            dfValue = NOTVALUE
        End If
        Call reloadComboBox(rcmVillage, Me.ComboBoxVillage, _
                            initValue:=Me.ComboBoxMO, defValue:=dfValue)
    End If
End Sub


Private Sub ComboBoxVillage_Change()
' ----------------------------------------------------------------------------
' обработка выбора населённого пункта
' Last update: 10.09.2017
' ----------------------------------------------------------------------------
    If Me.ComboBoxVillage.ListIndex > -1 Then
        Dim dfValue As Long
        Me.InfoPages.Visible = False
        Me.ComboBoxBldn.Clear
        If enableEvents Then
            dfValue = curItem.street.Id
        Else
            dfValue = NOTVALUE
        End If
        Call reloadComboBox(rcmStreet, Me.ComboBoxStreet, _
                            initValue:=Me.ComboBoxVillage, defValue:=dfValue)
    End If
End Sub


Private Sub ComboBoxStreet_Change()
' ----------------------------------------------------------------------------
' обработка выбора улицы
' Last update: 28.09.2016
' ----------------------------------------------------------------------------
    If Me.ComboBoxStreet.ListIndex > -1 Then
        Dim dfValue As Long
        If enableEvents Then
            dfValue = curItem.Id
        Else
            dfValue = NOTVALUE
        End If
        Call reloadComboBox(rcmListBldnNoId, Me.ComboBoxBldn, _
                            initValue:=Me.ComboBoxStreet, defValue:=dfValue)
    End If
End Sub


Private Sub ComboBoxBldn_Change()
' ----------------------------------------------------------------------------
' обработка выбора дома
' 20.08.2021
' ----------------------------------------------------------------------------
' при изменении дома показывается и заполняется информация по нему
    If Me.ComboBoxBldn.ListIndex > -1 Then
        If Not enableEvents Then
            Call terminateVars
            formBldnId = CLng(Me.ComboBoxBldn)
            Call initialCurItem
        End If
        ' отображение multipage с информацией
        Me.InfoPages.Visible = True
        ' флаг необходимости обновления списка работ
        workChanged = True
        planWorkChanged = True
        oldWorkChanged = True
        updateFlats = ls_clean
        ' заполнение активной вкладки
        If Me.InfoPages.Value = selPage Then
            Call loadInfo(selPage)
        Else
            Me.InfoPages.Value = selPage
        End If
        
        updateExpenses = True
        Call reloadComboBox(rcmBldnExpenseTerms, Me.ComboBoxExpenseTerms, _
                                    initValue:=curItem.Id, addAllItems:=True)
    Else
        ' если поле "номер дома" пусто, то прячем multipage
        If Not enableEvents Then Me.LabelCurItem.Caption = ""
        Me.InfoPages.Visible = False
    End If
End Sub


Private Sub InfoPages_Change()
' ----------------------------------------------------------------------------
' изменение вкладки
' Last update: 03.07.2016
' ----------------------------------------------------------------------------
    Set curWork = Nothing
    Call loadInfo(Me.InfoPages.Value)
    selPage = Me.InfoPages.Value
End Sub


Private Sub loadInfo(pageNum As InfoPagesEnum)
' ----------------------------------------------------------------------------
' вызов процедуры заполнения нужной вкладки
' 26.11.2021
' ----------------------------------------------------------------------------
    If curItem.Id = NOTVALUE Then Exit Sub
    Select Case pageNum
        Case InfoPagesEnum.ipInfo
            Call loadCommonInfo
            Call loadServiceInfo
            Call loadDogovorInfo
        Case InfoPagesEnum.ipTechInfo
            Call loadTechInfo
        Case InfoPagesEnum.ipLand
            Call loadLandInfo
        Case InfoPagesEnum.ipWorks
            Call reloadWorksYearCombo
        Case InfoPagesEnum.ipPlanWorks
            Call loadPlanWorksInfo
        Case InfoPagesEnum.ipOldWorks
            Call loadOldWorksInfo
        Case InfoPagesEnum.ipExpense
            Call loadExpenseInfo
        Case InfoPagesEnum.ipFlats
            Call loadFlatsInfo
        Case InfoPagesEnum.ipOffersWorks
            Call loadOffersWorks
        Case InfoPagesEnum.ipCommonPropertyElements
            Call loadCommonPropertyElements
    End Select
End Sub


Private Sub LabelCurItem_Click()
' ----------------------------------------------------------------------------
' по щелчку на адресе - показ истории субсчетов
' 02.08.2021
' ----------------------------------------------------------------------------
    If itemInitial Then
        Call showSubaccountHistory
    End If  ' itemInitial
End Sub


Private Sub loadCommonInfo()
' ----------------------------------------------------------------------------
' заполнение общей информации
' Last update: 09.04.2019
' ----------------------------------------------------------------------------
    Me.TextBoxBldnCadastral = curItem.CadastralNo
    Me.CheckBoxDisRepair = curItem.DisRepair
    Call selectComboBoxValue(Me.ComboBoxImprovement, _
                                                    curItem.Improvement.Id)
    Call selectComboBoxValue(Me.ComboBoxBldnType, curItem.BldnType)
    Call selectComboBoxValue(Me.ComboBoxEnergoClass, curItem.EnergoClass.Id)
    Me.TextBoxSiteNo.Value = curItem.SiteNo
    Me.TextBoxFias.Value = curItem.Fias
    Me.TextBoxGisGuid.Value = curItem.GisGuid
    ' код дома
    Me.LabelBldnIdValue.Caption = curItem.Id
    
    Call changeCommonState(False)
End Sub


Private Sub loadDogovorInfo()
' ----------------------------------------------------------------------------
' заполнение информации о договоре
' Last update: 25.05.2018
' ----------------------------------------------------------------------------
    Call selectComboBoxValue(Me.ComboBoxMCChange, curItem.uk.Id)
    Call selectComboBoxValue(Me.ComboBoxContractorChange, _
                                                        curItem.Contractor.Id)
    Call selectComboBoxValue(Me.ComboBoxDogovor, curItem.Dogovor.Id)
    If curItem.Contractor.Id <> 0 Then
        Call selectComboBoxValue(Me.ComboBoxManHourMode, _
                curItem.ManHourCost.Mode.Id)
    Else
        Call selectComboBoxValue(Me.ComboBoxManHourMode, 0)
    End If
    Me.LabelManHourMode.Caption = "Режим человекочаса (" & _
            curItem.ManHourCost.Cost & " руб.)"
    
    Me.CheckBoxOutReport.Value = curItem.outReport
    Me.CheckBoxOutReport.Caption = format(Me.CheckBoxOutReport, "Yes/No")
    Me.LabelContractInfo.Caption = "Договор №" & curItem.ContractNo & _
                                                " от " & curItem.ContractDate
    Call changeDogovorInfoState(False)
End Sub


Private Sub loadServiceInfo()
' ----------------------------------------------------------------------------
' заполнение информации об услугах
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        Call selectComboBoxValue(Me.ComboBoxColdWater, curItem.ColdWater.Id)
        Call selectComboBoxValue(Me.ComboBoxHotWater, curItem.HotWater.Id)
        Call selectComboBoxValue(Me.ComboBoxGas, curItem.Gas.Id)
        Call selectComboBoxValue(Me.ComboBoxHeating, curItem.Heating.Id)
        Call changeServiceState(False)
        Call reloadServicesList
    End If
End Sub


Public Sub reloadServicesList()
' ----------------------------------------------------------------------------
' заполнение услуг
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    Dim curList As New bldn_services
    Dim i As Long, j As Long
        
    Me.ListViewServices.Visible = False
    Me.ListViewServices.Visible = True
    Me.ListViewServices.ListItems.Clear
    ' заполнение списка работ
    curList.initial curItem.Id
            
    If curList.count > 0 Then
        ' заполнение данными
        Dim listX As listItem
        For i = 1 To curList.count
            Set listX = Me.ListViewServices.ListItems.add(, , _
                                                        curList(i).Service.Id)
            For j = 1 To FormBldnServices.bsMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormBldnServices.bsInputsCount).text = _
                                            curList(i).inputsCount
            listX.ListSubItems(FormBldnServices.bsModeId).text = _
                                            curList(i).Mode.Id
            listX.ListSubItems(FormBldnServices.bsServiceName).text = _
                                            curList(i).Service.Name
            listX.ListSubItems(FormBldnServices.bsModeName).text = _
                                            curList(i).Mode.Name
            listX.ListSubItems(FormBldnServices.bsPossibleCounter).text = _
                                            BoolToYesNo(curList(i).canCounter)
            listX.ListSubItems(FormBldnServices.bsNote).text = _
                                            curList(i).Note
        Next i
        Set listX = Nothing

        Call AppNewAutosizeColumns(Me.ListViewServices)
        Me.ListViewServices.ColumnHeaders(FormBldnServices.bsServiceId + 1).Width = 0
        Me.ListViewServices.ColumnHeaders(FormBldnServices.bsModeId + 1).Width = 0
    End If
    Set curList = Nothing
End Sub


Private Sub ButtonAddService_Click()
' ----------------------------------------------------------------------------
' Обработка кнопки добавления услуги
' Last update: 20.08.2018
' ----------------------------------------------------------------------------
    BldnServicesForm.BldnId = curItem.Id
    BldnServicesForm.Show
    Call reloadServicesList
End Sub


Private Sub ButtonChangeService_Click()
' ----------------------------------------------------------------------------
' Обработка кнопки изменения услуги
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    With BldnServicesForm
        .serviceId = Me.ListViewServices.selectedItem.text
        .BldnId = curItem.Id
        .Show
    End With
    Call reloadServicesList
End Sub


Private Sub ButtonServiceDelete_Click()
' ----------------------------------------------------------------------------
' Обработка кнопки удаления услуги
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    Dim ans As Boolean
    Dim tmp As bldn_service
    ans = ConfirmDeletion(Me.ListViewServices.selectedItem.ListSubItems( _
                                        FormBldnServices.bsServiceName).text)
    If ans Then
        Set tmp = New bldn_service
        tmp.initial curItem.Id, Me.ListViewServices.selectedItem.text
        tmp.delete
        Call reloadServicesList
    End If
    Set tmp = Nothing
End Sub


Private Sub changeCommonState(newState As Boolean)
' ----------------------------------------------------------------------------
' доступность элементов секции с общей информацией
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    Dim curCont As Control
    
    For Each curCont In Me.FrameTotalInfo.Controls
        If Not TypeName(curCont) = "CommandButton" Then
            curCont.Enabled = newState
        End If
    Next curCont
    Me.TextBoxSiteNo.Enabled = CurrentUser.isAdmin And newState
    Me.ButtonTotalChange.Caption = IIf(newState, SAVE_CAPTION, CHANGE_CAPTION)
    Me.ButtonTotalCancel.Visible = newState
End Sub


Private Sub changeDogovorInfoState(newState As Boolean)
' ----------------------------------------------------------------------------
' доступность элементов секции с информацией о договоре
' 30.05.2022
' ----------------------------------------------------------------------------
    Dim curCont As Control
    
    For Each curCont In Me.FrameManageInfo.Controls
        If Not TypeName(curCont) = "CommandButton" Then
            curCont.Enabled = newState
        End If
    Next curCont
    If curItem.Contractor.Id = 0 Then Me.ComboBoxManHourMode.Enabled = False
    Me.BtnDogovorChange.Caption = IIf(newState, SAVE_CAPTION, CHANGE_CAPTION)
    Me.BtnDogovorCancel.Visible = newState
End Sub


Private Sub changeServiceState(newState As Boolean)
' ----------------------------------------------------------------------------
' доступность элементов секции с услугами
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    Dim curCont As Control
    
    For Each curCont In Me.FrameService.Controls
        If Not TypeName(curCont) = "CommandButton" Then
            curCont.Enabled = newState
        End If
    Next curCont
    Me.BtnServiceCancel.Visible = newState
    Me.BtnServiceChange.Caption = IIf(newState, SAVE_CAPTION, CHANGE_CAPTION)
End Sub


Private Sub CheckBoxOutReport_Change()
' ----------------------------------------------------------------------------
' изменение подписи флажка вывода отчёта
' Last update: 08.12.2016
' ----------------------------------------------------------------------------
    Me.CheckBoxOutReport.Caption = format(Me.CheckBoxOutReport, "Yes/no")
End Sub


Private Sub BtnAnalisysElectro_Click()
' ----------------------------------------------------------------------------
' Вывод анализа ИПУ электроэнергии
' Last update: 01.06.2021
' ----------------------------------------------------------------------------
    Me.Hide
    Call AnalisysMeters(curItem.Id, curItem.AddressWOTown)
    Unload Me
End Sub


Private Sub CheckBoxBlock_Change()
' ----------------------------------------------------------------------------
' изменение подписи флажка блокировки
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    Me.CheckBoxBlock.Caption = format(Me.CheckBoxBlock, "Yes/No")
End Sub


Private Sub ButtonTotalChange_Click()
' ----------------------------------------------------------------------------
' обработка нажатия на кнопку изменения/сохранения общей информации
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    If Me.ButtonTotalChange.Caption = CHANGE_CAPTION Then
        ' изменение
        Call changeCommonState(True)
    Else
        ' сохранение
        Call updateCommonInfo
        Call loadCommonInfo
    End If
End Sub


Private Sub ButtonTotalCancel_Click()
' ----------------------------------------------------------------------------
' обработка кнопки отмены изменений общей информации
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    Call loadCommonInfo
End Sub


Private Sub updateCommonInfo()
' ----------------------------------------------------------------------------
' изменение общей информации
' Last update: 26.09.2018
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    curItem.updateCommon newImprovement:=Me.ComboBoxImprovement.Value, _
                        newBldnType:=Me.ComboBoxBldnType.Value, _
                        newCadastral:=Me.TextBoxBldnCadastral.Value, _
                        newDisRepair:=Me.CheckBoxDisRepair.Value, _
                        newEnergoClass:=Me.ComboBoxEnergoClass.Value, _
                        newSiteNo:=Me.TextBoxSiteNo.Value, _
                        newFias:=Me.TextBoxFias.Value, _
                        newGisGuid:=Me.TextBoxGisGuid.Value
    GoTo cleanHandler

errHandler:
    MsgBox "Не удалось сохранить изменения:" & vbCr & Err.Description, _
                                            vbExclamation, "Ошибка сохранения"
    GoTo cleanHandler

cleanHandler:
    On Error GoTo 0
End Sub


Private Sub BtnDogovorChange_Click()
' ----------------------------------------------------------------------------
' нажатие на кнопку изменения/сохранения информации о договоре
' Last update: 27.06.2017
' ----------------------------------------------------------------------------
    If Me.BtnDogovorChange.Caption = CHANGE_CAPTION Then
        ' изменение
        Call changeDogovorInfoState(True)
    Else
        ' сохранение
        Call updateDogovorInfo
        Call loadDogovorInfo
    End If
End Sub


Private Sub BtnDogovorCancel_Click()
' ----------------------------------------------------------------------------
' обработка кнопки отмены изменений информации о договоре
' Last update: 27.06.2017
' ----------------------------------------------------------------------------
    Call loadDogovorInfo
End Sub


Private Sub updateDogovorInfo()
' ----------------------------------------------------------------------------
' изменение информации о договоре
' Last update: 30.05.2022
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    curItem.updateDogovor newMC:=Me.ComboBoxMCChange.Value, _
                        newContractor:=Me.ComboBoxContractorChange.Value, _
                        newDogovor:=Me.ComboBoxDogovor.Value, _
                        newOutReport:=Me.CheckBoxOutReport.Value, _
                        newManHourMode:=Me.ComboBoxManHourMode.Value
    GoTo cleanHandler

errHandler:
    MsgBox "Не удалось сохранить изменения:" & vbCr & Err.Description, _
                                            vbExclamation, "Ошибка сохранения"
    GoTo cleanHandler

cleanHandler:
    On Error GoTo 0
End Sub


Private Sub BtnServiceChange_Click()
' ----------------------------------------------------------------------------
' обработка нажатия на кнопку изменения/сохранения услуг
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    If Me.BtnServiceChange.Caption = CHANGE_CAPTION Then
        ' изменение услуг, активируются все необходимые поля
        Call changeServiceState(True)
    Else
        ' сохранение изменений
        Call updateServiceInfo
        Call loadServiceInfo
    End If
End Sub


Private Sub BtnServiceCancel_Click()
' ----------------------------------------------------------------------------
' обработка кнопки отмены изменений услуг
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Call loadServiceInfo
End Sub


Private Sub updateServiceInfo()
' ----------------------------------------------------------------------------
' изменение информации об услугах
' Last update: 05.04.2018
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    curItem.updateServices newHeating:=Me.ComboBoxHeating.Value, _
                            newHotWater:=Me.ComboBoxHotWater.Value, _
                            newColdWater:=Me.ComboBoxColdWater.Value, _
                            newGas:=Me.ComboBoxGas.Value
    GoTo cleanHandler

errHandler:
    MsgBox "Не удалось сохранить изменения:" & vbCr & Err.Description, _
                                            vbExclamation, "Ошибка сохранения"
    GoTo cleanHandler

cleanHandler:
    On Error GoTo 0
End Sub


Private Sub ButtonDeleteBldn_Click()
' ----------------------------------------------------------------------------
' удалить дом
' Last update: 08.12.2016
' ----------------------------------------------------------------------------
    If Not ConfirmDeletion("этот дом") Then
        GoTo cleanHandler
    End If
                                                
    On Error GoTo errHandler
    curItem.delete
    MsgBox "Дом успешно удалён"
    Me.ComboBoxBldn.ListIndex = -1
    Call terminateVars
    Call reloadComboBox(rcmListBldnNoId, Me.ComboBoxBldn, _
                                        initValue:=Me.ComboBoxStreet.Value)
    GoTo cleanHandler

errHandler:
    MsgBox "Ошибка удаления дома" & vbCrLf & Err.Description, vbExclamation

cleanHandler:
End Sub


Private Sub BtnPassport_Click()
' ----------------------------------------------------------------------------
' создать паспорт дома
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    If getPrefetchWork = SERVICE_GLOBAL_TYPE Then
        Dim itId As Long
        itId = curItem.Id
        Unload Me
        RunReport1Form ("passport:" & itId)
    End If
End Sub


Private Sub ButtonCreatePrilCommonProp_Click()
' ----------------------------------------------------------------------------
' создать отчёт "характеристики общего имущества"
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim itId As Long
    Dim tmpHeader As String
    Dim CurCursor As XlMousePointer
    
    itId = curItem.Id
    
    TitleForm.Show
    tmpHeader = TitleForm.TextBoxTitle.Value
    Unload TitleForm
    
    If StrComp(tmpHeader, NOTSTRING) <> 0 Then
        CurCursor = Application.Cursor
        Application.Cursor = xlWait
        Call BldnCommonReport(itId, tmpHeader)
        Unload Me
        Application.Cursor = CurCursor
    End If
End Sub


Private Sub loadTechInfo()
' ----------------------------------------------------------------------------
' заполнение технической информации
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        If curTechInfo.BldnId = NOTVALUE Then
            Call initialCurTechInfo
        End If
        Me.TextFloorMax = curTechInfo.FloorMax
        Me.TextFloorMin = curTechInfo.FloorMin
        
        Me.TextYear = ""
        Me.CheckBoxNotBuiltYear = False
        If curTechInfo.BuiltYear = 0 Then
            Me.CheckBoxNotBuiltYear = True
        Else
            Me.TextYear = curTechInfo.BuiltYear
        End If
        Me.TextCommissioningYear = ""
        Me.CheckBoxNoCommissioningYear = False
        If curTechInfo.CommissioningYear = 0 Then
            Me.CheckBoxNoCommissioningYear = True
        Else
            Me.TextCommissioningYear = curTechInfo.CommissioningYear
        End If
        Me.TextEntrance = curTechInfo.EntranceCount
        Me.TextVault = curTechInfo.VaultsCount
        Me.TextDepreciation = curTechInfo.Depreciation
        Me.TextStairsSquare = curTechInfo.StairsSquare
        Me.TextCorridorSquare = curTechInfo.CorridorSquare
        Me.TextOtherMOPSquare = curTechInfo.OtherSquare
        Me.LabelMOPSquare = curTechInfo.MOPSquare
        Me.TextAtticSquare = curTechInfo.AtticSquare
        Me.TextVaultSquare = curTechInfo.VaultSquare
        Me.TextBoxStairs = curTechInfo.StairsCount
        Me.TextBoxStructuralVolume = curTechInfo.StructuralVolume
        Me.CheckBoxColdWater = curTechInfo.HasOdpuColdWater
        Me.CheckBoxCommon = curTechInfo.HasOdpuCommon
        Me.CheckBoxElectro = curTechInfo.HasOdpuElectro
        Me.CheckBoxHeating = curTechInfo.HasOdpuHeating
        Me.CheckBoxHotWater = curTechInfo.HasOdpuHotWater
        Me.CheckBoxHasDoorPhone = curTechInfo.HasDoorPhone
        Me.CheckBoxHasDoorPhone.Caption = IIf(Me.CheckBoxHasDoorPhone, _
                                            "Есть домофон", "Нет домофона")
        Me.CheckBoxHasDoorCloser = curTechInfo.HasDoorCloser
        Me.CheckBoxHasDoorCloser.Caption = IIf(Me.CheckBoxHasDoorCloser, _
                                            "Есть доводчики", "Нет доводчиков")
        Me.TextBoxDoorPhoneComment = curTechInfo.DoorPhoneComment
        Me.CheckBoxThermoregulator = curTechInfo.HasThermoregulator
        Me.TextBoxSquareBanisters = curTechInfo.ProcessedSurface.SquareBanisters
        Me.TextBoxSquareDoorHandles = curTechInfo.ProcessedSurface.SquareDoorHandles
        Me.TextBoxSquareDoors = curTechInfo.ProcessedSurface.SquareDoors
        Me.TextBoxSquareWindowSills = curTechInfo.ProcessedSurface.SquareWindowSills
        Me.TextBoxSquareMailBoxes = curTechInfo.ProcessedSurface.SquareMailBoxes
        Me.TextBoxSquareRadiators = curTechInfo.ProcessedSurface.SquareRadiators
        Call selectComboBoxValue(Me.ComboBoxWallMaterial, _
                                curTechInfo.WallMaterial.Id)
        
        Call reloadTmpCounters
        Call changeTechInfoState(False)
        
    End If
End Sub


Private Sub reloadTmpCounters()
' ----------------------------------------------------------------------------
' заполнение списка приборов учёта
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    On Error Resume Next
    Dim curList As New tmp_counters
    curList.initialByBldn curItem.Id
    curList.fillListView Me.ListViewAct
    Set curList = Nothing
End Sub


Private Sub BtnAddCounter_Click()
' ----------------------------------------------------------------------------
' Добавление прибора учёта
' Last update: 15.09.2020
' ----------------------------------------------------------------------------
    Call RunTmpCounterForm(curItem.Address, curItem.Id)
    Call reloadTmpCounters
End Sub


Private Sub BtnChangeCounter_Click()
' ----------------------------------------------------------------------------
' Изменение прибора учёта
' Last update: 15.09.2020
' ----------------------------------------------------------------------------
    If Me.ListViewAct.selectedItem Then
        Call RunTmpCounterForm(curItem.Address, curItem.Id, _
                                            Me.ListViewAct.selectedItem, True)
        Call reloadTmpCounters
    End If
End Sub


Private Sub BtnDeleteCounter_Click()
' ----------------------------------------------------------------------------
' Удаление прибора учёта
' Last update: 15.09.2020
' ----------------------------------------------------------------------------
    If Me.ListViewAct.selectedItem Then
        On Error GoTo errHandler
        Dim tmpItem As New tmp_counter
        tmpItem.initial (Me.ListViewAct.selectedItem)
        If ConfirmDeletion(tmpItem.Name) Then
            tmpItem.delete
            Call reloadTmpCounters
        End If
errHandler:
        If Err.Number <> 0 Then
            Dim errMsg As String
            If errorHasNoPrivilegies(Err.Description) Then
                errMsg = "У Вас нет прав на удаление прибора учёта"
            Else
                errMsg = Err.Description
            End If
            Err.Clear
            MsgBox "Ошибка удаления: " & vbCr & errMsg, _
                                vbOKOnly + vbExclamation, "Ошибка удаления"
        End If
        Set tmpItem = Nothing
    End If

End Sub


Private Sub BtnTechChange_Click()
' ----------------------------------------------------------------------------
' обработка нажатия на кнопку изменения/сохранения тех.информации
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    If Me.BtnTechChange.Caption = CHANGE_CAPTION Then
        Call changeTechInfoState(True)
    ElseIf Me.BtnTechChange.Caption = SAVE_CAPTION Then
        Call updateTechInfo
        Call loadTechInfo
    End If
End Sub


Private Sub BtnTechCancel_Click()
' ----------------------------------------------------------------------------
' обработка кнопки отмены изменений технической информации
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Call loadTechInfo
End Sub


Private Sub CheckBoxNotBuiltYear_Change()
' ----------------------------------------------------------------------------
' галочка установлен ли год постройки МКД
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    Me.TextYear.Enabled = Not Me.CheckBoxNotBuiltYear
    If Not Me.CheckBoxNotBuiltYear Then Me.TextYear.SetFocus
End Sub


Private Sub CheckBoxNoCommissioningYear_Change()
' ----------------------------------------------------------------------------
' галочка установлен ли год ввода в эксплуатацию МКД
' Last update: 17.05.2016
' ----------------------------------------------------------------------------
    Me.TextCommissioningYear.Enabled = Not Me.CheckBoxNoCommissioningYear
    If Not Me.CheckBoxNoCommissioningYear Then _
                                        Me.TextCommissioningYear.SetFocus
End Sub


Private Sub CheckBoxHasDoorPhone_Click()
' ----------------------------------------------------------------------------
' изменение подписи флажка наличия домофона
' Last update: 09.11.2020
' ----------------------------------------------------------------------------
    Me.CheckBoxHasDoorPhone.Caption = IIf(Me.CheckBoxHasDoorPhone, _
                                            "Есть домофон", "Нет домофона")
End Sub


Private Sub CheckBoxHasDoorCloser_Click()
' ----------------------------------------------------------------------------
' изменение подписи флажка наличия доводчика
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    Me.CheckBoxHasDoorCloser.Caption = IIf(Me.CheckBoxHasDoorCloser, _
                                        "Есть доводчики", "Нет доводчиков")
End Sub


Private Sub updateTechInfo()
' ----------------------------------------------------------------------------
' изменение технической информации
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    Dim bYear As Long, cYear As Long
    
    If Me.CheckBoxNotBuiltYear Then
        bYear = 0
    Else
        bYear = longValue(Me.TextYear)
    End If
    If Me.CheckBoxNoCommissioningYear Then
        cYear = 0
    Else
        cYear = longValue(Me.TextCommissioningYear)
    End If
    
    curTechInfo.update newFloorMax:=longValue(Me.TextFloorMax), newFloorMin:=longValue(Me.TextFloorMin), _
                        newVaultsCount:=longValue(Me.TextVault), newEntranceCount:=longValue(Me.TextEntrance), _
                        newStairsCount:=longValue(Me.TextBoxStairs), newDepreciation:=dblValue(Me.TextDepreciation), _
                        newCorridorSquare:=dblValue(Me.TextCorridorSquare), newStairsSquare:=dblValue(Me.TextStairsSquare), _
                        newOtherSquare:=dblValue(Me.TextOtherMOPSquare), newAtticSquare:=dblValue(Me.TextAtticSquare), _
                        newVaultSquare:=dblValue(Me.TextVaultSquare), newWallmaterial:=dblValue(Me.ComboBoxWallMaterial), _
                        newBuiltYear:=bYear, newStructuralVolume:=dblValue(Me.TextBoxStructuralVolume), _
                        newCommissioningYear:=cYear, newHasHotWater:=Me.CheckBoxHotWater.Value, _
                        newHasColdWater:=Me.CheckBoxColdWater.Value, newHasCommon:=Me.CheckBoxCommon.Value, _
                        newHasHeating:=Me.CheckBoxHeating.Value, newHasElectro:=Me.CheckBoxElectro.Value, _
                        newHasDoorPhone:=Me.CheckBoxHasDoorPhone, newDoorPhoneComment:=Trim(Me.TextBoxDoorPhoneComment), _
                        newHasThermoregulator:=Me.CheckBoxThermoregulator, _
                        newSquareBanisters:=Me.TextBoxSquareBanisters, newSquareDoors:=Me.TextBoxSquareDoors, _
                        newSquareWindowSills:=Me.TextBoxSquareWindowSills, newSquareDoorHandles:=Me.TextBoxSquareDoorHandles, _
                        newSquareMailBoxes:=Me.TextBoxSquareMailBoxes, newSquareRadiators:=Me.TextBoxSquareRadiators, _
                        newHasDoorCloser:=Me.CheckBoxHasDoorCloser
    
    GoTo cleanHandler
    
errHandler:
    MsgBox "Не удалось сохранить изменения:" & vbCr & Err.Description, _
                                            vbExclamation, "Ошибка сохранения"
    GoTo cleanHandler

cleanHandler:
    On Error GoTo 0
End Sub


Private Sub changeTechInfoState(newState As Boolean)
' ----------------------------------------------------------------------------
' доступность элементов секции с услугами
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Dim curControl As Control
    
    For Each curControl In Me.FrameTechInfo.Controls
        If TypeName(curControl) = "TextBox" Or _
                            TypeName(curControl) = "CheckBox" Then
            curControl.Enabled = newState
        End If
    Next curControl
    Me.ComboBoxWallMaterial.Enabled = newState
    ' год постройки
    If Me.CheckBoxNotBuiltYear Then Me.TextYear.Value = ""
    ' год ввода в эксплуатацию
    If Me.CheckBoxNoCommissioningYear Then _
                                Me.TextCommissioningYear.Value = ""
    Me.TextYear.Enabled = _
                    Not Me.CheckBoxNotBuiltYear.Value And newState
    Me.TextCommissioningYear.Enabled = _
                    Not Me.CheckBoxNoCommissioningYear And newState
    
    Me.BtnTechChange.Caption = IIf(newState, SAVE_CAPTION, CHANGE_CAPTION)
    Me.BtnTechCancel.Visible = newState
End Sub


Private Sub loadLandInfo()
' ----------------------------------------------------------------------------
' заполнение информации о З/У
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        If curLandInfo.BldnId = NOTVALUE Then
            Call initialCurLandInfo
        End If
        Me.TBCadastralNo = curLandInfo.CadastralNo
        Me.TBInventoryLandArea = curLandInfo.InventoryArea
        Me.TBUseLandArea = curLandInfo.UseArea
        Me.TBBuiltUpArea = curLandInfo.BuiltUp
        Me.TBUndevelopedArea = curLandInfo.Undeveloped
        Me.TBHardCoatings = curLandInfo.HardCoatings
        Me.TBDriveWays = curLandInfo.DriveWays
        Me.TBSideWalks = curLandInfo.SideWalks
        Me.TBOthers = curLandInfo.Others
        Me.TBSurveySquare = curLandInfo.SurveyArea
        Me.CheckBoxFences = curLandInfo.Fences
        Me.CheckBoxFences.Caption = "Ограждения (" & _
                                BoolToYesNo(curLandInfo.Fences) & ")"
        Me.CheckBoxSAF = curLandInfo.SAF
        Me.CheckBoxSAF.Caption = "Малые архитектурные формы (" & _
                                BoolToYesNo(curLandInfo.SAF) & ")"
        Me.TextBoxBenches = curLandInfo.Benches
        Call changeLandInfoState(False)
    End If
End Sub


Private Sub changeLandInfoState(newState As Boolean)
' ----------------------------------------------------------------------------
' доступность элементов секции с информацией о З/У
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Dim curControl As Control
    
    ' доступность всех текстовых полей и CheckBox
    For Each curControl In Me.InfoPages("PageLandArea").Controls
        If StrComp(TypeName(curControl), "TextBox", vbTextCompare) = 0 Or _
            StrComp(TypeName(curControl), "CheckBox", vbTextCompare) = 0 Then
            curControl.Enabled = newState
        End If
    Next curControl
    Me.BtnLandChange.Caption = IIf(newState, SAVE_CAPTION, CHANGE_CAPTION)
    Me.BtnLandCancel.Visible = newState
End Sub


Private Sub BtnLandChange_Click()
' ----------------------------------------------------------------------------
' обработка кнопки изменения/сохранения информации о З/У
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    If Me.BtnLandChange.Caption = CHANGE_CAPTION Then
        Call changeLandInfoState(True)
    ElseIf Me.BtnLandChange.Caption = SAVE_CAPTION Then
        Call updateLandInfo
        Call loadLandInfo
    End If
End Sub


Private Sub BtnLandCancel_Click()
' ----------------------------------------------------------------------------
' обработка кнопки отмены изменения земельного участка
' Last update: 11.05.2016
' ----------------------------------------------------------------------------
    Call loadLandInfo
End Sub


Private Sub updateLandInfo()
' ----------------------------------------------------------------------------
' изменение информации о З/У
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    curLandInfo.update newCadastral:=Me.TBCadastralNo, _
                        newInventory:=dblValue(Me.TBInventoryLandArea), _
                        newUse:=dblValue(Me.TBUseLandArea), _
                        newBuilt:=dblValue(Me.TBBuiltUpArea), _
                        newUndeveloped:=dblValue(Me.TBUndevelopedArea), _
                        newHard:=dblValue(Me.TBHardCoatings), _
                        newDrive:=dblValue(Me.TBDriveWays), _
                        newSide:=dblValue(Me.TBSideWalks), _
                        newOther:=dblValue(Me.TBOthers), _
                        newSurvey:=dblValue(Me.TBSurveySquare), _
                        newSAF:=Me.CheckBoxSAF.Value, _
                        newFences:=Me.CheckBoxFences, _
                        newBenches:=Me.TextBoxBenches
    
    GoTo cleanHandler
    
errHandler:
    MsgBox "Не удалось сохранить изменения:" & vbCr & Err.Description, _
                                            vbExclamation, "Ошибка сохранения"
    GoTo cleanHandler

cleanHandler:
    On Error GoTo 0
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.listItem)
' ----------------------------------------------------------------------------
' при выборе работы инициализациия переменной
' Last update: 22.04.2016
' ----------------------------------------------------------------------------
    Set curWork = New work_class
    curWork.initial CLng(Item)
End Sub


Private Sub ComboBoxGWT_Change()
' ----------------------------------------------------------------------------
' обработка выбора типа ремонта
' Last update: 18.10.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxGWT.ListIndex > -1 Then
        workChanged = True
        Call reloadWorksYearCombo
    End If
End Sub


Private Sub ComboBoxWorksYearSelect_Change()
' ----------------------------------------------------------------------------
' обработка выбора года работ
' Last update: 12.10.2016
' ----------------------------------------------------------------------------
    If Me.ComboBoxWorksYearSelect.ListIndex > -1 Then
        workChanged = True
        Call reloadWorkList
    End If
End Sub


Private Sub BtnAddWork_Click()
' ----------------------------------------------------------------------------
' добавление работы. форма прячется
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    If formPrefGWT = SERVICE_GLOBAL_TYPE Then
        With WorkMaintenanceForm
            Set .curBldn = curItem
            Me.Hide
            .Show 'vbModeless
        End With
    Else
        WorkForm.BldnId = curItem.Id
        WorkForm.formMcId = curItem.uk.Id
        WorkForm.formPrefGWT = formPrefGWT
        WorkForm.formContractorId = curItem.Contractor.Id
        WorkForm.LabelAddress.Caption = curItem.Address
        Me.Hide
        WorkForm.Show vbModeless
    End If
End Sub


Private Sub BtnChangeWork_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки изменения работы
' Last update: 01.10.2019
' ----------------------------------------------------------------------------
    If Not curWork Is Nothing Then
        If curWork.Id <> NOTVALUE Then
            If formPrefGWT = SERVICE_GLOBAL_TYPE Then
                Dim curMWork As New work_maintenance
                On Error Resume Next
                curMWork.initialByWorkId curWork.Id
                On Error GoTo 0
                If curMWork.Id = NOTVALUE Then
                    Err.Clear
                    GoTo changeCommonWork
                End If
                
                With WorkMaintenanceForm
                    Set .curBldn = curItem
                    Set .curWork = curMWork
                    Me.Hide
                    workChanged = True
                    .Show vbModeless
                End With
            Else
changeCommonWork:
                Me.Hide
                workChanged = True
                With WorkForm
                    Set .changedWork = curWork
                    .LabelAddress.Caption = curItem.Address
                    .BldnId = curItem.Id
                    .Show vbModeless
                End With
            End If
        End If
    End If
End Sub


Private Sub BtnDelWork_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки удаления работы
' Last update: 26.09.2018
' ----------------------------------------------------------------------------
    If Not curWork Is Nothing Then
        If curWork.Id > 0 Then
            ' запрос подтверждения
            Dim deleteString As String
            deleteString = curWork.WorkKind.Name & " от " & _
                            curWork.StringDate & " на сумму " & curWork.sum
            If ConfirmDeletion(deleteString) Then
                On Error GoTo errHandler
                curWork.delete
                workChanged = True
                Call loadInfo(selPage)
            End If
        End If
    End If
    GoTo cleanHandler

errHandler:
    Call setWorkErrMsg(Err.Description)
cleanHandler:
End Sub


Private Sub BtnWorkReport_Click()
' ----------------------------------------------------------------------------
' обработка кнопки отчёта работ по дому
' Last update: 27.05.2019
' ----------------------------------------------------------------------------
    Call RunReport1Form("bldnwork:" & curItem.Id)
    Unload Me
End Sub


Private Sub reloadWorksYearCombo()
' ----------------------------------------------------------------------------
' заполнение списка годов
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim tmp As Collection
    
    If Not workChanged Then Exit Sub
    If IsNull(Me.ComboBoxGWT.Value) Then Exit Sub
    Set tmp = worksYears(Me.ComboBoxGWT.Value, curItem.Id)
    Me.ComboBoxWorksYearSelect.Clear
    Me.ComboBoxWorksYearSelect.AddItem "Все"
    For i = 1 To tmp.count
        Me.ComboBoxWorksYearSelect.AddItem tmp(i)
    Next i
        
    For i = 0 To Me.ComboBoxWorksYearSelect.ListCount - 1
        If Me.ComboBoxWorksYearSelect.list(i) = CStr(Year(Date)) Then
            Me.ComboBoxWorksYearSelect.ListIndex = i
            Exit For
        End If
    Next i
    ' если в текущем году работ не было, то выбирается пункт "Все"
    If Me.ComboBoxWorksYearSelect.ListIndex = -1 Then
        Me.ComboBoxWorksYearSelect.ListIndex = 0
    End If
    workChanged = False
End Sub


Private Sub reloadWorkList()
' ----------------------------------------------------------------------------
' заполнение списка работ
' Last update: 27.05.2019
' ----------------------------------------------------------------------------
    Dim curList As New bldnworks
    Dim curWorkYear As Long
    Dim beginDate As Long, EndDate As Long
    Dim SUStatus As Boolean
    Dim i As Long, j As Long
        
    Set curWork = Nothing
    
    If Me.ComboBoxWorksYearSelect.ListIndex <= 0 Then
        beginDate = ALLVALUES
        EndDate = ALLVALUES
    Else
        curWorkYear = CLng(Me.ComboBoxWorksYearSelect.Value)
        beginDate = terms.FirstTermInYear(curWorkYear).Id
        EndDate = terms.LastTermInYear(curWorkYear).Id
    End If
    
    Me.ListViewList.Visible = False
    
    Call setWorkErrMsg("")
    Me.ListViewList.ListItems.Clear
    ' заполнение списка работ
    curList.initialByBldn ItemId:=curItem.Id, gwtId:=CLng(Me.ComboBoxGWT), _
                            beginDate:=beginDate, EndDate:=EndDate
            
    If curList.count > 0 Then
        ' заполнение данными
        Dim listX As listItem
        For i = 1 To curList.count
            Set listX = Me.ListViewList.ListItems.add(, , curList(i).Id)
            For j = 1 To FormWorkListEnum.fwlMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormWorkListEnum.fwlContractor).text = _
                                                curList(i).cName
            listX.ListSubItems(FormWorkListEnum.fwlDate).text = _
                                                dateToStr(curList(i).wDate)
            listX.ListSubItems(FormWorkListEnum.fwlDogovor).text = _
                                                curList(i).wDogovor
            listX.ListSubItems(FormWorkListEnum.fwlFSource).text = _
                                                curList(i).wFSource
            listX.ListSubItems(FormWorkListEnum.fwlNote).text = _
                                                curList(i).wNote
            listX.ListSubItems(FormWorkListEnum.fwlPrintFlag).text = _
                                    format(curList(i).wPrintFlag, "Yes/No")
            listX.ListSubItems(FormWorkListEnum.fwlSI).text = _
                                                            curList(i).wSI
            listX.ListSubItems(FormWorkListEnum.fwlSum).text = _
                                                            curList(i).wSum
            listX.ListSubItems(FormWorkListEnum.fwlVolume).text = _
                                                            curList(i).wVolume
            listX.ListSubItems(FormWorkListEnum.fwlWK).text = _
                                                            curList(i).wkName
            listX.ListSubItems(FormWorkListEnum.fwlWT).text = _
                                                            curList(i).wtName
        Next i
        Set listX = Nothing

        Call AppNewAutosizeColumns(Me.ListViewList)
            
    End If
    Me.ListViewList.Visible = True
    
    Me.ListViewList.Visible = False
    Me.ListViewList.Visible = True
    
    Me.LabelSumWork = curList.TotalSum
    Set curList = Nothing
End Sub


Private Sub loadPlanWorksInfo()
' ----------------------------------------------------------------------------
' заполнение вкладки с плановыми работами
' 25.04.2022
' ----------------------------------------------------------------------------
    Me.CommandButtonDeletePlan.Visible = CurrentUser.isAdmin
    Me.LabelPlanSubAccount.Caption = "Плановые поступления в месяц: " & _
            curItem.SubaccountPlanSum & ", собираемость:" & _
            curItem.SubaccountPercent & "%" & _
            vbCrLf & _
            "План на конец года: " & curItem.SubaccountPlanEndSum & " (" & _
            curItem.SubaccountPlanEndWithPercentSum & ")"
                            
    Call reloadPlanWorkYearsList
End Sub


Private Sub ComboBoxPlanYears_Change()
' ----------------------------------------------------------------------------
' при выборе года заполнение плановых работ этого года
' Last update: 30.07.2019
' ----------------------------------------------------------------------------
    If Me.ComboBoxPlanYears.ListIndex > -1 Then
        planWorkChanged = True
        Call reloadPlanWorkList
    End If
End Sub


Private Sub reloadPlanWorkYearsList()
' ----------------------------------------------------------------------------
' заполнение списка годов планируемых работ
' Last update: 30.07.2019
' ----------------------------------------------------------------------------
    Dim curList As Collection
    Dim i As Long

    If Not planWorkChanged Then Exit Sub
    
    planWorkChanged = False
    With Me.ComboBoxPlanYears
        Set curList = DBgetPlanWorksYears(curItem.Id)
        .Clear
        .AddItem "Все"
        For i = 1 To curList.count
            .AddItem CStr(curList(i))
        Next i
        .ListIndex = 0
    End With
    planWorkChanged = True

End Sub


Private Sub reloadPlanWorkList()
' ----------------------------------------------------------------------------
' заполнение списка планируемых работ
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    Dim curList As New plan_work_list
    Dim curPlanWork As plan_work_class
    Dim i As Long, j As Long
    Dim listX As listItem
    Dim bDate As Date, eDate As Date
    
    If Not planWorkChanged Then Exit Sub
    
    If Me.ComboBoxPlanYears.ListIndex = 0 Then
        bDate = NOTDATE
        eDate = NOTDATE
    Else
        bDate = DateSerial(Me.ComboBoxPlanYears.Value, 1, 1)
        eDate = DateSerial(Year(bDate), 12, 31)
    End If
    curList.initialByBldn curItem.Id, bDate, eDate
        
    Me.ListViewPlanWork.Visible = False
    Me.ListViewPlanWork.ListItems.Clear
    
    If curList.count > 0 Then
        ' заполнение данными
        For i = 1 To curList.count
            Set curPlanWork = curList(i)
            Set listX = Me.ListViewPlanWork.ListItems.add(, , curPlanWork.Id)
            For j = 1 To FormPlanWorkEnum.fpwMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormPlanWorkEnum.fpwWK).text = _
                                        curPlanWork.WorkKind.Name
            listX.ListSubItems(FormPlanWorkEnum.fpwContractor).text = _
                                        curPlanWork.Contractor.Name
            listX.ListSubItems(FormPlanWorkEnum.fpwDate).text = _
                                        curPlanWork.StringDate
            listX.ListSubItems(FormPlanWorkEnum.fpwGWT).text = _
                                        curPlanWork.GWT.Name
            listX.ListSubItems(FormPlanWorkEnum.fpwMC).text = _
                                        curPlanWork.MC.Name
            listX.ListSubItems(FormPlanWorkEnum.fpwNote).text = _
                                        curPlanWork.Note
            listX.ListSubItems(FormPlanWorkEnum.fpwPrivateNote).text = _
                                        curPlanWork.PrivateNote
            listX.ListSubItems(FormPlanWorkEnum.fpwStatus).text = _
                                        curPlanWork.Status.Name
            listX.ListSubItems(FormPlanWorkEnum.fpwSum).text = curPlanWork.sum
            listX.ListSubItems(FormPlanWorkEnum.fpwEmployee).text = _
                                        curPlanWork.Employee
            listX.ListSubItems(FormPlanWorkEnum.fpwWorkRef).text = _
                                        curPlanWork.workRef
            listX.ListSubItems(FormPlanWorkEnum.fpwBeginDate).text = _
                                        IIf(curPlanWork.beginDate = NOTDATE, _
                                            "", curPlanWork.beginDate)
            listX.ListSubItems(FormPlanWorkEnum.fpwEndDate).text = _
                                        IIf(curPlanWork.EndDate = NOTDATE, _
                                            "", curPlanWork.EndDate)
            listX.ListSubItems(FormPlanWorkEnum.fpwSmetaSum).text = _
                                        curPlanWork.smetaSum
            If curPlanWork.Status.inPlan Then
                If Year(curPlanWork.WorkDate) < Year(Now) Then
                    highlightListItem listX, vbRed
                ElseIf Month(curPlanWork.WorkDate) < Month(Now) Then
                    highlightListItem listX, vbBlue
                End If      ' if year() < now
            End If      ' if inPlan
        Next i
        Set listX = Nothing
        
        Call AppNewAutosizeColumns(Me.ListViewPlanWork)
            
    End If
    Me.ListViewPlanWork.Visible = True
    
    ' эти две строки лечат убегание ListView из своего положения
    Me.ListViewPlanWork.Visible = False
    Me.ListViewPlanWork.Visible = True
    
    planWorkChanged = False
    Set curList = Nothing
    Set curPlanWork = Nothing
End Sub


Private Sub ListViewPlanWork_ItemClick(ByVal Item As MSComctlLib.listItem)
' ----------------------------------------------------------------------------
' выбор планируемой работы
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    On Error Resume Next
    Set curPlanWork = New plan_work_class
    curPlanWork.initial CLng(Item)
End Sub


Private Sub CommandButtonAddPlan_Click()
' ----------------------------------------------------------------------------
' нажатие кнопки добавления плановой работы
' Last update: 16.02.2018
' ----------------------------------------------------------------------------
    Me.Hide
    PlanWorkForm.BldnId = curItem.Id
    PlanWorkForm.mcId = curItem.uk.Id
    PlanWorkForm.LabelBldn.Caption = curItem.Address
    PlanWorkForm.Show
End Sub


Private Sub CommandButtonChangePlan_Click()
' ----------------------------------------------------------------------------
' нажатие кнопки изменения плановой работы
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    If Not curPlanWork Is Nothing Then
        If curPlanWork.Id <> NOTVALUE Then
            PlanWorkForm.BldnId = curItem.Id
            PlanWorkForm.mcId = curItem.uk.Id
            PlanWorkForm.workId = curPlanWork.Id
            PlanWorkForm.LabelBldn.Caption = curItem.Address
            Set curPlanWork = Nothing
            Me.Hide
            PlanWorkForm.Show
        End If
    End If
End Sub


Private Sub CommandButtonDeletePlan_Click()
' ----------------------------------------------------------------------------
' нажатие кнопки удаления плановой работы
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    If Not curPlanWork Is Nothing Then
        If curPlanWork.Id <> NOTVALUE Then
            If ConfirmDeletion(curPlanWork.WorkKind.Name & " " & _
                                                curPlanWork.StringDate) Then
                curPlanWork.delete
                planWorkChanged = True
                Call loadInfo(ipPlanWorks)
            End If
        End If
    End If
End Sub


Private Sub CommandButtonPrintPlanList_Click()
' ----------------------------------------------------------------------------
' распечатка план-листа для плановой работы
' Last update: 14.11.2019
' ----------------------------------------------------------------------------
    If Not curPlanWork Is Nothing Then
        If curPlanWork.Status.inPlan Then
            Call ReportBldnPlanList(curItem, curPlanWork)
            Unload Me
        End If
    End If
End Sub


Private Sub showErrorMessage(message As String)
' ----------------------------------------------------------------------------
' вывод сообщения об ошибке
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    MsgBox message, vbOKOnly + vbExclamation, "Ошибка"
End Sub


Private Sub loadOldWorksInfo()
' ----------------------------------------------------------------------------
' заполнение вкладки с старыми работами
' Last update: 05.05.2018
' ----------------------------------------------------------------------------
    Call reloadOldWorkList
End Sub


Private Sub reloadOldWorkList()
' ----------------------------------------------------------------------------
' заполнение списка старых работ
' Last update: 05.05.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As listItem
    
    If Not oldWorkChanged Then Exit Sub
    
    Set oldWorks = New old_works
    oldWorks.initial curItem.Id
    
    Me.ListViewOldWorks.ListItems.Clear
    
    If oldWorks.this.count > 0 Then
        ' заполнение данными
        For i = 1 To oldWorks.this.count
            Set curOldWork = oldWorks.this(i)
            Set listX = Me.ListViewOldWorks.ListItems.add(, , curOldWork.Id)
            For j = 1 To FormOldWorksEnum.fowMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormOldWorksEnum.fowName).text = curOldWork.wName
            listX.ListSubItems(FormOldWorksEnum.fowNote).text = curOldWork.wNote
            listX.ListSubItems(FormOldWorksEnum.fowOBF).text = BoolToYesNo(curOldWork.wOBF)
            listX.ListSubItems(FormOldWorksEnum.fowOBN).text = curOldWork.wOBN
            listX.ListSubItems(FormOldWorksEnum.fowSum).text = curOldWork.wSum
            listX.ListSubItems(FormOldWorksEnum.fowVolume).text = curOldWork.wVolume
            listX.ListSubItems(FormOldWorksEnum.fowYear).text = curOldWork.wYear
        Next i
        Set listX = Nothing
        
        Call AppNewAutosizeColumns(Me.ListViewOldWorks)
            
    End If
    
    oldWorkChanged = False
    Set curOldWork = Nothing
End Sub


Private Sub ListViewOldWorks_ItemClick(ByVal Item As MSComctlLib.listItem)
' ----------------------------------------------------------------------------
' выбор старой работы
' Last update: 05.05.2018
' ----------------------------------------------------------------------------
    Set curOldWork = oldWorks.this(CStr(Item))
End Sub


Private Sub ButtonDeleteOldWork_Click()
' ----------------------------------------------------------------------------
' нажатие кнопки удаления старой работы
' Last update: 05.05.2018
' ----------------------------------------------------------------------------
    If Not curOldWork Is Nothing Then
        If curOldWork.Id <> NOTVALUE Then
            If ConfirmDeletion(curOldWork.wName & " " & _
                                                curOldWork.wYear) Then
                curOldWork.delete
                oldWorkChanged = True
                Call loadInfo(ipOldWorks)
            End If
        End If
    End If
End Sub

Private Sub loadExpenseInfo()
' ----------------------------------------------------------------------------
' заполнение вкладки "Цена"
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    If updateExpenses Then
        Me.ComboBoxExpenseNames.ListIndex = 0
    End If
    Call reloadExpenseList
    If Not CurrentUser.isAdmin Then Me.BtnChangeExpense.Visible = False
End Sub


Public Sub reloadExpenseList()
' ----------------------------------------------------------------------------
' заполнение структуры
' Last update: 05.07.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As listItem
    Dim tmpExp As bldn_expenses
    Dim isLast As Boolean
    Dim expId As Long, TermId As Long
    
    expId = NOTVALUE
    TermId = NOTVALUE
    If Me.ComboBoxExpenseNames.ListIndex > 0 Then
        expId = Me.ComboBoxExpenseNames.Value
        Me.ComboBoxExpenseTerms.ListIndex = 0
    ElseIf Me.ComboBoxExpenseTerms.ListIndex > 0 Then
        TermId = Me.ComboBoxExpenseTerms.Value
        Me.ComboBoxExpenseNames.ListIndex = 0
    End If
    
    If expId = NOTVALUE And TermId = NOTVALUE Then
        isLast = True
    Else
        isLast = False
    End If
    
    Set tmpExp = New bldn_expenses
    tmpExp.initial curItem.Id, TermId:=TermId, expenseId:=expId, _
                                                    lastMonthExpenses:=isLast
    With Me.ListViewExpenses
        .Visible = False
        .ListItems.Clear
        
        For i = 1 To tmpExp.count
            Set listX = .ListItems.add(, , tmpExp(i).Id)
            For j = 1 To FormBldnLastExpenses.fbleMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormBldnLastExpenses.fbleName).text = tmpExp(i).ExpenseItem.Name
            listX.ListSubItems(FormBldnLastExpenses.fbleDate).text = tmpExp(i).Term.Name
            listX.ListSubItems(FormBldnLastExpenses.fbleBldnName).text = _
                                                                tmpExp(i).Name
            listX.ListSubItems(FormBldnLastExpenses.fblePrice).text = _
                                                                tmpExp(i).price
            listX.ListSubItems(FormBldnLastExpenses.fblePlanSum).text = _
                                                                tmpExp(i).planSum
            listX.ListSubItems(FormBldnLastExpenses.fbleFactSum).text = _
                                                                tmpExp(i).factSum
        Next i
        
        ' обновление подписи в label
        Me.LabelExpensesInfo.Caption = tmpExp.ExpensePrice & _
                ", плановые расходы: " & tmpExp.ExpensePlanSum
        
        Set listX = Nothing
        Set tmpExp = Nothing
        
        Call AppNewAutosizeColumns(Me.ListViewExpenses)
        .ColumnHeaders(1).Width = 0
        .Visible = True
        .Visible = False
        .Visible = True
        .SetFocus
    End With
    updateExpenses = False
End Sub


Private Sub ComboBoxExpenseNames_Change()
' ----------------------------------------------------------------------------
' обработка изменения статьи расходов
' Last update: 04.07.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxExpenseNames.ListIndex > -1 And _
                                    Me.ComboBoxExpenseTerms.ListCount > 1 Then
        If Me.ComboBoxExpenseNames.ListIndex = 0 Then
            Me.ComboBoxExpenseTerms.ListIndex = 1
            Me.ComboBoxExpenseTerms.Enabled = True
            Me.LabelExpensesError.Caption = ""
        Else
            Me.ComboBoxExpenseTerms.ListIndex = 0
            Me.ComboBoxExpenseTerms.Enabled = False
            Me.LabelExpensesError.Caption = ""
        End If
    End If
End Sub


Private Sub ComboBoxExpenseTerms_Change()
' ----------------------------------------------------------------------------
' обработка изменения месяца статей расходов
' Last update: 04.07.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxExpenseTerms.ListIndex = 0 And _
                                Me.ComboBoxExpenseNames.ListIndex = 0 And _
                                Me.ComboBoxExpenseTerms.ListCount > 1 Then
        Me.LabelExpensesError.Caption = _
                    "Необходимо выбрать либо период, либо статью расходов"
    Else
        Me.LabelExpensesError.Caption = ""
    End If
End Sub


Private Sub BtnShowExpenses_Click()
' ----------------------------------------------------------------------------
' показ структуры с выбранными параметрами
' Last update: 05.07.2018
' ----------------------------------------------------------------------------
    updateExpenses = True
    Call reloadExpenseList
End Sub


Private Sub ListViewExpenses_KeyPress(KeyAscii As Integer)
' ----------------------------------------------------------------------------
' при нажатии Enter на структуре - редактирование
' Last update: 06.07.2018
' ----------------------------------------------------------------------------
    If KeyAscii = 13 And getUserId() = 1 Then
        Call editExpense(Me.ListViewExpenses.selectedItem)
    End If
End Sub


Private Sub ListViewExpenses_DblClick()
' ----------------------------------------------------------------------------
' при двойном щелчке на структуре показывать окно её изменения
' Last update: 06.07.2018
' ----------------------------------------------------------------------------
    If CurrentUser.isAdmin Then Call editExpense(Me.ListViewExpenses.selectedItem)
End Sub


Private Sub BtnChangeExpense_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки изменения структуры
' Last update: 06.07.2018
' ----------------------------------------------------------------------------
    Call editExpense(Me.ListViewExpenses.selectedItem)
End Sub


Private Sub editExpense(curExpense As listItem)
' ----------------------------------------------------------------------------
' запуск окна изменения структуры
' Last update: 10.04.2019
' ----------------------------------------------------------------------------
    Load ChangeExpenseForm
    With ChangeExpenseForm
        If curExpense Is Nothing Then Exit Sub
        .expenseId = curExpense
        .LabelDescription.Caption = curItem.Address & vbCr & _
            curExpense.ListSubItems(FormBldnLastExpenses.fbleName).text & _
            vbCr & curExpense.ListSubItems(FormBldnLastExpenses.fbleDate).text
        .TextBoxPrice.Value = curExpense.ListSubItems( _
                                        FormBldnLastExpenses.fblePrice).text
        .TextBoxPlanSum.Value = curExpense.ListSubItems( _
                                            FormBldnLastExpenses.fblePlanSum).text
        .TextBoxFactSum.Value = curExpense.ListSubItems( _
                                            FormBldnLastExpenses.fbleFactSum).text
    End With
    ChangeExpenseForm.Show
End Sub



Private Sub BtnMonthReport_Click()
' ----------------------------------------------------------------------------
' Вывод месячного акта выполнения работ
' 27.09.2022
' ----------------------------------------------------------------------------
    Me.Hide
    Call ReportBldnWorkCompletition(curItem.Id)
    Unload Me
End Sub


Private Sub BtnExpenseToGis_Click()
' ----------------------------------------------------------------------------
' Окно параметров для выгрузки плана работ в ГИС
' Last update: 20.05.2019
' ----------------------------------------------------------------------------
    If Len(curItem.GisGuid) = 0 Then
        MsgBox "Необходимо заполнить GUID дома в ГИС ЖКХ", vbInformation
        Exit Sub
    End If
    
    Report1Form.Tag = "gisexpense:" & curItem.Id
    Unload Me
    Report1Form.Show
End Sub


Private Sub loadFlatsInfo()
' ----------------------------------------------------------------------------
' заполнение информации по помещениям
' 20.08.2021
' ----------------------------------------------------------------------------
    If updateFlats = ls_clean Then
        Call reloadComboBox(rcmFlatTerms, Me.ComboBoxFlatTerms, curItem.Id)
        updateFlats = ls_data
        If Me.ComboBoxFlatTerms.ListCount > 0 Then
            Call fillFlatsListView
        End If
    End If
End Sub


Private Sub fillFlatsListView()
' ----------------------------------------------------------------------------
' заполнение списка квартир
' 01.09.2021
' ----------------------------------------------------------------------------
    If Not updateFlats = ls_data Then Exit Sub
    Dim curList As New flats
    curList.initialByBldnAndTerm curItem.Id, CLng(Me.ComboBoxFlatTerms.Value)
    curList.fillFullInfoListForm Me.ListViewFlats
    Me.LabelFlats = "Всего: " & curList.FlatsCount & _
                " (" & curList.FlatsSquare & " кв.м.)" & _
            " в т.ч. жилых: " & curList.ResidentalCount & _
                " (" & curList.ResidentalSquare & " кв.м.)" & _
            ", нежилых: " & curList.NonResidentalCount & _
                " (" & curList.NonResidentalSquare & " кв.м.)"
    Set curList = Nothing
End Sub


Private Sub ComboBoxFlatTerms_Change()
' ----------------------------------------------------------------------------
' при изменении месяца заполнение списка квартир
' 13.08.2021
' ----------------------------------------------------------------------------
    Call fillFlatsListView
End Sub


Private Sub ListViewFlats_DblClick()
' ----------------------------------------------------------------------------
' показ истории квартиры при двойном клике
' 22.06.2022
' ----------------------------------------------------------------------------
    Call FlatAccruedHistory(Me.ListViewFlats.selectedItem)
End Sub


Private Sub BtnAccruedHistory_Click()
' ----------------------------------------------------------------------------
' показ истории начислений по квартире
' 22.06.2022
' ----------------------------------------------------------------------------
    Call FlatAccruedHistory(Me.ListViewFlats.selectedItem)
End Sub


Private Sub FlatAccruedHistory(selectedItem As listItem)
' ----------------------------------------------------------------------------
' история начислений по квартире
' 22.06.2022
' ----------------------------------------------------------------------------
    With ListForm
        .setParameter "flatId", selectedItem
        .setParameter "HideButtons", True
        .setParameter "flatNo", selectedItem.ListSubItems(1)
        .formType = lftFlatAccrueds
        .Show
    End With
End Sub


Private Sub BtnFlatHistory_Click()
' ----------------------------------------------------------------------------
' истории квартиры при двойном клике
' 20.06.2022
' ----------------------------------------------------------------------------
    With ListForm
        .setParameter "flatId", Me.ListViewFlats.selectedItem
        .setParameter "HideButtons", True
        .formType = lftFlatHistory
        .Show
    End With
End Sub


Private Sub BtnAddSignature_Click()
' ----------------------------------------------------------------------------
' добавление подписи коменданта
' 19.10.2022
' ----------------------------------------------------------------------------
    Call showChairmanSignHistory
End Sub


Private Sub loadOffersWorks()
' ----------------------------------------------------------------------------
' заполнение информации по предложениям к договорам
' 14.10.2021
' ----------------------------------------------------------------------------
    Call fillOffersWorksListView
End Sub


Private Sub fillOffersWorksListView()
' ----------------------------------------------------------------------------
' заполнение списка предложенных работ
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim curList As New offer_works
    curList.reload curItem.Id
    curList.fillListform Me.ListViewOffersWorks
    Set curList = Nothing
End Sub


Private Sub BtnPrintOffersWorks_Click()
' ----------------------------------------------------------------------------
' экспорт предложений по работам
' 15.10.2021
' ----------------------------------------------------------------------------
    Dim offers As New offer_works
    offers.ExportList curItem
    Unload Me
    Set offers = Nothing
End Sub


Private Sub ButtonAddOfferWork_Click()
' ----------------------------------------------------------------------------
' добавление предложений по работам
' 21.10.2021
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim tmpItem As New offer_work
    tmpItem.add NOTVALUE, curItem.Id, DateSerial(Year(Now) + 1, 1, 1), _
                NOTSTRING, NOTVALUE, NOTVALUE
    tmpItem.showForm False
    Call loadOffersWorks

errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Ошибка"
    Set tmpItem = Nothing
End Sub


Private Sub ButtonChangeOffersWork_Click()
' ----------------------------------------------------------------------------
' изменение предложений по работам
' 21.10.2021
' ----------------------------------------------------------------------------
    If Me.ListViewOffersWorks.selectedItem Is Nothing Then Exit Sub
    
    On Error GoTo errHandler
    
    Dim tmpItem As New offer_work
    tmpItem.initial Me.ListViewOffersWorks.selectedItem
    tmpItem.showForm True
    Call loadOffersWorks

errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Ошибка"
    Set tmpItem = Nothing

End Sub


Private Sub ButtonDelete_Click()
' ----------------------------------------------------------------------------
' удаление предложений по работам
' 21.10.2021
' ----------------------------------------------------------------------------
    If Me.ListViewOffersWorks.selectedItem Is Nothing Then Exit Sub
    
    On Error GoTo errHandler
    
    Dim tmpItem As New offer_work
    tmpItem.initial Me.ListViewOffersWorks.selectedItem
    If ConfirmDeletion(tmpItem.Name) Then
        tmpItem.delete
        Call loadOffersWorks
    End If

errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Ошибка"
    Set tmpItem = Nothing

End Sub


Private Sub loadCommonPropertyElements()
' ----------------------------------------------------------------------------
' Заполнение информации об элементах общего имущества
' 25.05.2022
' ----------------------------------------------------------------------------

    Dim tmpList As New bldn_common_properties
    tmpList.reload curItem.Id, ShowAll:=True
    tmpList.fillListView Me.ListViewCPE
    Set tmpList = Nothing
    Me.ButtonChangeCPE.Enabled = False
    Me.ButtonChangeElementContain.Enabled = False
End Sub


Private Sub ListViewCPE_ItemClick(ByVal Item As MSComctlLib.listItem)
' ----------------------------------------------------------------------------
' Отображение нужных кнопок при выборе элемента общего имущества
' 12.04.2022
' ----------------------------------------------------------------------------
    Me.ButtonChangeCPE.Enabled = False
    Me.ButtonChangeElementContain.Enabled = False
    
    Dim tmpItem As New bldn_common_property
    tmpItem.initialByListViewRow Item
    If tmpItem.IsElement Then Me.ButtonChangeElementContain.Enabled = True
    If (tmpItem.IsElement And tmpItem.m_IsUsing) Or _
            (tmpItem.IsParameter And tmpItem.m_IsUsing) Then
        Me.ButtonChangeCPE.Enabled = True
    End If
    Set tmpItem = Nothing
End Sub


Private Sub ListViewCPE_DblClick()
' ----------------------------------------------------------------------------
' Обработка двойного клика на элементе списка
' 12.04.2022
' ----------------------------------------------------------------------------
    Call changeCPE
End Sub


Private Sub ButtonChangeElementContain_Click()
' ----------------------------------------------------------------------------
' Изменение наличия элемента общего имущества в доме
' 30.05.2022
' ----------------------------------------------------------------------------

    On Error GoTo errHandler
    Dim tmpItem As New bldn_common_property
    tmpItem.initialByListViewRow Me.ListViewCPE.selectedItem
        
    Dim ans As Integer, qText As String
    If tmpItem.m_IsUsing Then
        qText = "Удалить элемент " & tmpItem.m_Name & " из дома?" & vbCr & _
                "При этом будут удалены все значения параметров!"
    Else
        qText = "Поставить наличие элемента " & tmpItem.m_Name & "?"
    End If
    ans = MsgBox(qText, vbYesNo + vbExclamation, "Внимание")
    If ans = vbYes Then
        Dim bcpItem As New bldn_common_property_element
        Dim tmpLItem As Object, i As Long
        Set tmpLItem = Me.ListViewCPE.selectedItem
        
        bcpItem.add tmpItem.m_ElementId, curItem.Id, tmpItem.m_Name, tmpItem.m_State, tmpItem.m_IsUsing
        bcpItem.changeContain (Not bcpItem.IsContain)
        Set bcpItem = Nothing
        Call loadCommonPropertyElements
        
        For i = 1 To Me.ListViewCPE.ListItems.count
            If Me.ListViewCPE.ListItems(i) = tmpLItem Then
                Me.ListViewCPE.ListItems(i).EnsureVisible
                Application.Wait Now() + TimeValue("0:00:00")
                Exit For
            End If
        Next i
    End If
    Set tmpItem = Nothing

errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Ошибка"
End Sub


Private Sub ButtonChangeCPE_Click()
' ----------------------------------------------------------------------------
' Кнопка изменения состояния элемента общего имущества в доме
' 12.04.2022
' ----------------------------------------------------------------------------
    Call changeCPE
End Sub


Private Sub ButtonInspectionReport_Click()
' ----------------------------------------------------------------------------
' Кнопка печати акта обследования элементов общего имущества
' 05.05.2022
' ----------------------------------------------------------------------------
    Me.Hide
    Call BldnInspectionReport(curItem)
    Unload Me
End Sub


Private Sub BtnBldnCPComposition_Click()
' ----------------------------------------------------------------------------
' Кнопка печати состава общего имущества и его состояния
' 25.05.2022
' ----------------------------------------------------------------------------
    Me.Hide
    Call BldnCompositionCommonProperties(curItem)
    Unload Me
End Sub


Private Sub changeCPE()
' ----------------------------------------------------------------------------
' Изменение состояния элемента общего имущества в доме
' 27.05.2022
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim rTmpItem As New bldn_common_property
    rTmpItem.initialByListViewRow Me.ListViewCPE.selectedItem
    
    Dim curLItem As Object
    Set curLItem = Me.ListViewCPE.selectedItem
    
    Dim tmpItem As Object
    If rTmpItem.IsElement Then
        Set tmpItem = New bldn_common_property_element
        tmpItem.add rTmpItem.m_ElementId, curItem.Id, rTmpItem.m_Name, rTmpItem.m_State, rTmpItem.m_IsUsing
        tmpItem.showForm (True)
    ElseIf rTmpItem.IsParameter Then
        Set tmpItem = New bldn_common_property_parameter
        tmpItem.add rTmpItem.m_ParameterId, curItem.Id, rTmpItem.m_Name, "", rTmpItem.m_State
        tmpItem.showForm True
    End If
    If Not rTmpItem.IsGroup Then
        Call loadCommonPropertyElements
    End If
    
    Dim i As Long
    For i = 1 To Me.ListViewCPE.ListItems.count
        If Me.ListViewCPE.ListItems(i) = curLItem Then
            Me.ListViewCPE.ListItems(i).EnsureVisible
            Exit For
        End If
    Next i
    Application.Wait (Now + TimeValue("0:00:01"))
    Me.ListViewCPE.SetFocus
    Me.ListViewCPE.ListItems(i).EnsureVisible
    Set curLItem = Nothing
    
errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Ошибка"
End Sub


Private Sub initialCurItem()
' ----------------------------------------------------------------------------
' инициализация внутренней переменной
' Last update: 09.04.2019
' ----------------------------------------------------------------------------
    Set curItem = New building_class
    If formBldnId <> NOTVALUE And formBldnId <> 0 Then
        curItem.initial formBldnId
        formBldnId = NOTVALUE
        Me.LabelCurItem = curItem.Address & ". Остаток на субсчете на " & _
                        curItem.SubaccountDate & ": " & curItem.SubaccountSum
    End If
End Sub


Private Sub initialCurTechInfo()
' ----------------------------------------------------------------------------
' инициализация внутренней переменной технической информации
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Set curTechInfo = New bldnTechInfo
    If curItem.Id <> NOTVALUE And curItem.Id <> 0 Then
        curTechInfo.initial curItem.Id
    End If
End Sub


Private Sub initialCurLandInfo()
' ----------------------------------------------------------------------------
' инициализация внутренней переменной информации о З/У
' Last update: 11.10.2016
' ----------------------------------------------------------------------------
    Set curLandInfo = New bldnLandInfo
    If curItem.Id <> NOTVALUE And curItem.Id <> 0 Then
        curLandInfo.initial curItem.Id
    End If
End Sub


Private Sub terminateVars()
' ----------------------------------------------------------------------------
' уничтожение всех переменных
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Set curItem = Nothing
    Set curLandInfo = Nothing
    Set curTechInfo = Nothing
    Set curWork = Nothing
    Set curPlanWork = Nothing
End Sub


Private Sub setWorkErrMsg(errMsg As String)
' ----------------------------------------------------------------------------
' сообщение на странице работ
' Last update: 26.09.2018
' ----------------------------------------------------------------------------
    Me.LabelWorksMsg.Caption = errMsg
    Me.LabelWorksMsg.ForeColor = RGB(255, 0, 0)
End Sub


Private Function itemInitial() As Boolean
' ----------------------------------------------------------------------------
' проверка на инициализацию объекта
' 02.08.2021
' ----------------------------------------------------------------------------
    itemInitial = False
    If Not curItem Is Nothing Then
        If curItem.Id <> NOTVALUE Then
            itemInitial = True
        End If
    End If
End Function


Private Function showSubaccountHistory()
' ----------------------------------------------------------------------------
' показать форму с историей по сусбчетам
' 17.08.2021
' ----------------------------------------------------------------------------
    With ListForm
        .setParameter "bldnId", curItem.Id
        .setParameter "HideButtons", True
        .formType = lftBldnSubaccounts
        .Show
    End With
End Function


Private Function showChairmanSignHistory()
' ----------------------------------------------------------------------------
' показать форму работы с подписями комендантов
' 19.10.2022
' ----------------------------------------------------------------------------
    With ListForm
        .setParameter "bldnId", curItem.Id
        .setParameter "HideButtonChange", True
        .formType = lftChairmansSign
        .Show
    End With
End Function
