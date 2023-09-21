Attribute VB_Name = "WorkModule"
Option Explicit
Option Private Module

Sub test()
    ReportForm.Tag = "site"
    ReportForm.show
End Sub

Sub createWorksheet()
' ----------------------------------------------------------------------------
' пересоздание листа со списком домов
' Last update: 07.05.2017
' ----------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim daStatus As Integer
    Dim i As Integer, j As Long
    Dim queryString As String
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
        
    With ThisWorkbook
        Set ws = .Worksheets(shtnmMain)
        ws.Unprotect shtPass
        
        ws.UsedRange.Clear
        i = 1
        ws.Cells(i, MainSheetEnum.msBldnNo).Value = "Дом"
        ws.Cells(i, MainSheetEnum.msCode).Value = "Код"
        ws.Cells(i, MainSheetEnum.msContractor).Value = "Подрядчик"
        ws.Cells(i, MainSheetEnum.msMD).Value = "МО"
        ws.Cells(i, MainSheetEnum.msStreet).Value = "Улица"
        ws.Cells(i, MainSheetEnum.msUK).Value = "УК"
        ws.Cells(i, MainSheetEnum.msVillage).Value = "Населенный пункт"
        ws.Cells(i, MainSheetEnum.msDogovor).Value = "Вид договора"
        ws.Cells(i, MainSheetEnum.msOutReport).Value = "Вывод"
        
        Set cmd = New ADODB.Command
        cmd.CommandText = "create_sheet"
        cmd.CommandType = adCmdStoredProc
        cmd.ActiveConnection = DBConnection.Connection
        Set rst = cmd.Execute
        
        If rst.EOF And rst.BOF Then Exit Sub
                        
        ' заполнение
        i = i + 1
        Do While Not rst.EOF
            ws.Cells(i, MainSheetEnum.msBldnNo).Value = rst!BldnNo
            ws.Cells(i, MainSheetEnum.msCode).Value = rst!bId
            ws.Cells(i, MainSheetEnum.msContractor).Value = rst!contname
            ws.Cells(i, MainSheetEnum.msDogovor).Value = rst!dogname
            ws.Cells(i, MainSheetEnum.msMD).Value = rst!mdName
            ws.Cells(i, MainSheetEnum.msOutReport).Value = BoolToYesNo(rst!outReport)
            ws.Cells(i, MainSheetEnum.msStreet).Value = rst!StreetName
            ws.Cells(i, MainSheetEnum.msUK).Value = rst!mcname
            ws.Cells(i, MainSheetEnum.msVillage).Value = rst!vilname
            i = i + 1
            rst.MoveNext
        Loop
        
        ws.UsedRange.Columns.AutoFit
        ws.UsedRange.AutoFilter
        ws.AutoFilter.ShowAllData
        ws.Protect shtPass, AllowFiltering:=True
        ws.Activate
        ws.Cells(2, 1).Select
        ActiveWindow.FreezePanes = True
    End With
End Sub


Function ConfirmDeletion(deletionValue As String) As Boolean
' ----------------------------------------------------------------------------
' подтверждение удаления
' Last update: 10.05.2016
' ----------------------------------------------------------------------------
    Dim delAnswer As Integer
    delAnswer = MsgBox("Вы действительно хотите удалить " & deletionValue & _
                                        "?", vbYesNo, "Подтвердите удаление")
    If delAnswer = vbYes Then
        ConfirmDeletion = True
    Else
        ConfirmDeletion = False
    End If
End Function


Function dblValue(strValue As Variant) As Double
' ----------------------------------------------------------------------------
' преобразование строки в double с заменой десятичного разделителя
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    dblValue = 0
    If IsNull(strValue) Then Exit Function
    strValue = Replace(strValue, ".", format(0, "."))
    strValue = Replace(strValue, ",", format(0, "."))
    If IsNumeric(strValue) Then dblValue = CDbl(strValue)
End Function


Function boolValue(strValue As Variant) As Boolean
' ----------------------------------------------------------------------------
' преобразование строки в boolean, NULL считается как False
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    boolValue = False
    If IsNull(strValue) Then Exit Function
    boolValue = strValue
End Function


Function longValue(strValue As Variant) As Long
' ----------------------------------------------------------------------------
' преобразование значения в Long
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    If IsNumeric(strValue) Then
        longValue = CLng(strValue)
    Else
        longValue = 0
    End If
End Function


Public Function DBgetString(str As Variant) As String
' ----------------------------------------------------------------------------
' Получение строки из базы, с преобразованием NULL в пустую строку
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    If IsNull(str) Then
        DBgetString = ""
    Else
        DBgetString = str
    End If
End Function


Public Function DBgetDate(str As Variant) As Date
' ----------------------------------------------------------------------------
' Получение строки из базы, с преобразованием NULL в 01/01/1900
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    If IsNull(str) Then
        DBgetDate = DateSerial(1900, 1, 1)
    Else
        DBgetDate = dateValue(str)
    End If
End Function


Public Function DBgetDateStr(str As Variant) As String
' ----------------------------------------------------------------------------
' Получение строки из базы, с преобразованием NULL в пустую строку
' Last update: 09.08.2018
' ----------------------------------------------------------------------------
    If IsNull(str) Then
        DBgetDateStr = ""
    Else
        DBgetDateStr = CStr(dateValue(str))
    End If
End Function


Function dateToStr(dateValue As Date) As String
' ----------------------------------------------------------------------------
' преобразование даты в строку
' Last update: 22.05.2016
' ----------------------------------------------------------------------------
    dateToStr = MonthName(Month(dateValue)) & " " & Year(dateValue)
End Function


Function BoolToYesNo(boolValue As Boolean, _
        Optional yesvalue As Integer = 0, _
        Optional trueString As String = NOTSTRING, _
        Optional falseString As String = NOTSTRING) As String
' ----------------------------------------------------------------------------
' преобразование булева значения в строку "ДА" (либо "Есть"), "НЕТ"
' с возможностью указать, какие строки должны отображаться
' 16.08.2021
' ----------------------------------------------------------------------------
    If trueString = NOTSTRING Then
        trueString = IIf(yesvalue = 0, "Да", "Есть")
    End If
    If falseString = NOTSTRING Then
        falseString = "Нет"
    End If
    BoolToYesNo = IIf(boolValue, trueString, falseString)
    
End Function


Function Int32ToBool(intValue As Integer) As Boolean
' ----------------------------------------------------------------------------
' преобразование числа в логическое (0-ложь, остальное-истина)
' Last update: 09.09.2016
' ----------------------------------------------------------------------------
    Int32ToBool = IIf(intValue = 0, False, True)
End Function


Sub workSheetSort(ws As Worksheet, colCount As Integer, colList As Variant)
' ----------------------------------------------------------------------------
' сортировка листа по указанным столбцам
' Last update: 10.05.2016
' ----------------------------------------------------------------------------
    Dim i As Integer
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    With ws.Sort
        .SortFields.Clear
        For i = 1 To colCount
            .SortFields.add Key:=.Parent.Columns(colList(i - 1)), _
                                SortOn:=xlSortOnValues, Order:=xlAscending, _
                                DataOption:=xlSortNormal
        Next i
        .SetRange .Parent.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = SUStatus
    
End Sub


Function getQueryString(fileName As String) As String
' ----------------------------------------------------------------------------
' считывание SQL-скрипта из файла
' Last update: 17.03.2021
' ----------------------------------------------------------------------------
    fileName = ThisWorkbook.Path & Application.PathSeparator & "sql" & _
                                        Application.PathSeparator & fileName
    getQueryString = getFileText(fileName)
    Debug.Print Len(getQueryString)
End Function


Function getTemplateString(fileName As String) As String
' ----------------------------------------------------------------------------
' формирование пути к шаблону отчета
' Last update: 22.05.2018
' ----------------------------------------------------------------------------
    Const templatesPath As String = "templates"
    
    getTemplateString = ThisWorkbook.Path & Application.PathSeparator & _
                        templatesPath & Application.PathSeparator & fileName
End Function


Function getThisPath() As String
' ----------------------------------------------------------------------------
' каталог до файла
' Last update: 27.06.2017
' ----------------------------------------------------------------------------
    getThisPath = ThisWorkbook.Path & Application.PathSeparator
End Function


Sub reloadComboBox(curType As ReloadComboMethods, _
                                    curCombo As Control, _
                                    Optional initValue As Long = NOTVALUE, _
                                    Optional defValue As Long = NOTVALUE, _
                                    Optional addAllItems As Boolean = False, _
                                    Optional addNotValue As Boolean = False, _
                                    Optional initValue2 As Long = NOTVALUE, _
                                    Optional initString As String = NOTSTRING)
' ----------------------------------------------------------------------------
' заполнение переданного списка, с последующим выбором
'      если это необходимо
' initValue - для инициализации списков по коду
' defValue - значение, показываемое по умолчанию
' addAllItems - нужно ли добавлять значение "---все---"
' initValue2 - второе значение, которое необходимо для инициализации списка
' initString - строковое значение, которое передаётся в необходимые списки
' params - словарь с параметрами
' 02.06.2022
' ----------------------------------------------------------------------------
    Dim curList As Object
    Dim i As Long
    Dim curListIndex As Integer
    
    On Error GoTo errHandler
    
    ' определение, какой список заполнять
    Select Case curType:
        Case ReloadComboMethods.rcmMd:
            Set curList = address_md_list
        Case ReloadComboMethods.rcmVillage:
            If initValue = NOTVALUE Then
                Set curList = address_village_list
            Else
                Set curList = New villages_in_md
                curList.initial initValue
            End If
        Case ReloadComboMethods.rcmStreet:
            Set curList = New address_street_list
            If initValue <> NOTVALUE Then curList.initial initValue
        Case ReloadComboMethods.rcmMC:
            Set curList = New uk_list
        Case ReloadComboMethods.rcmContractor:
            Set curList = contractor_list
        Case ReloadComboMethods.rcmGWT:
            Set curList = globalWorkType_list
        Case ReloadComboMethods.rcmWorkType:
            Set curList = worktype_list
        Case ReloadComboMethods.rcmWorkKind:
            Set curList = New workkind_list
            If initValue <> NOTVALUE Then curList.initial initValue
        Case ReloadComboMethods.rcmTerm:
            Set curList = terms
        Case ReloadComboMethods.rcmTermDESC:
            Set curList = New terms
            curList.reload False
        Case ReloadComboMethods.rcmListBldnNoId:
            Set curList = New bldn_no_id_list
            curList.initial initValue
        Case ReloadComboMethods.rcmListBldnAddressId:
            Set curList = New bldn_no_id_list
            curList.initialWithAddress
        Case ReloadComboMethods.rcmListBldnAddressIdByMD:
            Set curList = New bldn_no_id_list
            curList.initialWithAddress mdId:=initValue, dogovorType:=IIf(initValue2 = NOTVALUE, ALLVALUES, initValue2)
        Case ReloadComboMethods.rcmListManagedBldnAddressIdByMD:
            Set curList = New bldn_no_id_list
            curList.initialWithAddress mdId:=initValue, OnlyManaged:=True
        Case ReloadComboMethods.rcmListBldnAddressIdByStreet:
            Set curList = New bldn_no_id_list
            curList.initialWithAddress filterId:=initValue
        Case ReloadComboMethods.rcmListBldnAddressIdByVillage:
            Set curList = New bldn_no_id_list
            curList.initialWithAddress villageId:=initValue
        Case ReloadComboMethods.rcmImprovement:
            Set curList = improvement_list
        Case ReloadComboMethods.rcmDogovor:
            Set curList = dogovor_list
        Case ReloadComboMethods.rcmWallMaterial:
            Set curList = New wallmaterial_list
        Case ReloadComboMethods.rcmStreetTypes
            Set curList = street_types
        Case ReloadComboMethods.rcmMainContractor
            Set curList = New contractor_bldn_list
        Case ReloadComboMethods.rcmUsingMainContractor
            Set curList = New contractor_bldn_list
            curList.initial True
        Case ReloadComboMethods.rcmPlanStatuses
            Set curList = plan_statuses
        Case ReloadComboMethods.rcmPlanStatusesNewWork
            Set curList = plan_statuses_new_work
        Case ReloadComboMethods.rcmVillageTypes
            Set curList = village_types
        Case ReloadComboMethods.rcmGas
            Set curList = New id_name_list
            curList.initGas
        Case ReloadComboMethods.rcmHeating
            Set curList = New id_name_list
            curList.initHeating
        Case ReloadComboMethods.rcmHotWater
            Set curList = New id_name_list
            curList.initHotWater
        Case ReloadComboMethods.rcmColdWater
            Set curList = New id_name_list
            curList.initColdWater
        Case ReloadComboMethods.rcmEmployees
            Set curList = New employee_list
            curList.initial initValue
        Case ReloadComboMethods.rcmFSources
            Set curList = fsources
        Case ReloadComboMethods.rcmYesNo
            Set curList = New id_name_list
            If StrComp(initString, NOTSTRING) = 0 Then
                curList.initYesNo
            Else
                curList.initYesNoOther initString
            End If
        Case ReloadComboMethods.rcmServices
            Set curList = services
        Case ReloadComboMethods.rcmExpenseGroups
            Set curList = New expense_groups
        Case ReloadComboMethods.rcmExpenseItems
            Set curList = expense_items
        Case ReloadComboMethods.rcmBldnExpenseName
            Set curList = New id_name_list
            curList.initBldnExpenseName initValue, initValue2
        Case ReloadComboMethods.rcmBldnExpenseTerms
'            Set curList = New bldn_expense_terms
'            curList.reload initValue
            Set curList = New terms
            curList.loadBldnExpensesMonths (initValue)
            If curList.count > 0 Then defValue = curList(1).Id
        Case ReloadComboMethods.rcmServiceModes
            Set curList = New service_modes
            curList.reload initValue
        Case ReloadComboMethods.rcmUserRoles
            Set curList = New user_roles
            curList.reload
        Case ReloadComboMethods.rcmUsers
            Set curList = New users
        Case ReloadComboMethods.rcmUserHasRoles
            Set curList = New user_roles
            curList.reload initValue
        Case ReloadComboMethods.rcmUserHasNoRoles
            Set curList = New user_roles
            curList.reload initValue, False
        Case ReloadComboMethods.rcmAccessTypes
            Set curList = New id_name_list
            curList.initAccessTypes
        Case ReloadComboMethods.rcmBldnTypes
            Set curList = New id_name_list
            curList.initBldnTypes
        Case ReloadComboMethods.rcmEnergoClasses
            Set curList = energo_classes
        Case ReloadComboMethods.rcmRoleHasAccess
            Set curList = New id_name_list
            curList.initRoleAccess initValue, initValue2, True
        Case ReloadComboMethods.rcmRoleHasNoAccess
            Set curList = New id_name_list
            curList.initRoleAccess initValue, initValue2, False
        Case ReloadComboMethods.rcmPlanTerms:
            Set curList = New plan_terms
            curList.reload CDate(initValue), initValue2
            If defValue <> NOTVALUE Then defValue = curList.IdByDate(CDate(defValue))
        Case ReloadComboMethods.rcmWorkMaterialTypes:
            Set curList = material_types
        Case ReloadComboMethods.rcmRkcServices:
            Set curList = New rkc_services
        Case ReloadComboMethods.rcmUkServices:
            Set curList = uk_services
        Case ReloadComboMethods.rcmServiceTypes
            Set curList = New service_types
        Case ReloadComboMethods.rcmFlatTerms
            Set curList = New terms
            curList.loadFlatsMonths initValue
            If curList.count > 0 Then defValue = curList.LastTerm.Id
        Case ReloadComboMethods.rcmAddedTypes
            Set curList = New added_types
        Case ReloadComboMethods.rcmCommonPropertyGroup
            Set curList = common_property_groups
        Case ReloadComboMethods.rcmCommonPropertyElement
            Set curList = New common_property_elements
        Case ReloadComboMethods.rcmSubaccountTerms
            Set curList = New terms
            curList.loadSubAccountMonths
        Case ReloadComboMethods.rcmManHourModes
            Set curList = man_hour_cost_modes
        Case Else:
            Exit Sub
    End Select
    
    With curCombo
        .Clear
        .ColumnCount = 2
        .BoundColumn = ComboColumns.ccId + 1
        curListIndex = 0
        If addAllItems Then
            .AddItem
            .list(0, ccname) = ALL_STRING
            .list(0, ccId) = ALLVALUES
            curListIndex = curListIndex + 1
        End If
        If addNotValue Then
            .AddItem
            .list(0, ccname) = NOTSTRING
            .list(0, ccId) = NOTVALUE
            curListIndex = curListIndex + 1
        End If
        For i = 1 To curList.count
            .AddItem
            .list(curListIndex, ComboColumns.ccname) = curList(i).Name
            .list(curListIndex, ComboColumns.ccId) = curList(i).Id
            curListIndex = curListIndex + 1
        Next i
        .ColumnWidths = ";0"
        Set curList = Nothing
        
        ' при необходимости выбираем пункт
        If addAllItems And defValue = NOTVALUE Then defValue = ALLVALUES
        If defValue <> NOTVALUE Or addNotValue Then
            Call selectComboBoxValue(curCombo, defValue)
        End If
                
    End With
    GoTo cleanHandler
    
errHandler:
    If Not curCombo Is Nothing Then curCombo.Clear
    Err.Raise Err.Number, Err.Source, Err.Description
    
cleanHandler:
    Set curList = Nothing
'    Set curCombo = Nothing
End Sub


Sub selectComboBoxValue(curCB As Control, fValue As Long)
' ----------------------------------------------------------------------------
' выбор пункта из ComboBox
' Last update: 19.02.2018
' ----------------------------------------------------------------------------
    Dim i As Long
'    If fValue <> NOTVALUE Then
        With curCB
            For i = 0 To .ListCount - 1
                If CLng(.list(i, ComboColumns.ccId)) = fValue Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End With
'    End If
End Sub


Sub MoveListBoxElements(sourceLB As Object, destLB As Object, _
                                    Optional allElements As Boolean = False)
' ----------------------------------------------------------------------------
' перемещение элементов списка
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    If TypeName(sourceLB) = "ListBox" And TypeName(destLB) = "ListBox" Then
        Dim i As Long
        Dim selectedItems As New Collection
        If sourceLB.ListIndex > -1 Then
            ' сначала копирование
            For i = 0 To sourceLB.ListCount - 1
                If sourceLB.Selected(i) Or allElements Then
                    destLB.AddItem
                    destLB.list(destLB.ListCount - 1, ccname) = _
                                                    sourceLB.list(i, ccname)
                    destLB.list(destLB.ListCount - 1, ccId) = _
                                                    sourceLB.list(i, ccId)
                    selectedItems.add i
                End If
            Next i
            ' потом удаление
            For i = selectedItems.count To 1 Step -1
                sourceLB.RemoveItem selectedItems(i)
            Next i
        End If
        Set selectedItems = Nothing
    Else
        Err.Raise ERROR_OBJECT_NOT_VALID, "MoveListBoxElements", _
                                    "Неправильный объект, требуется ListBox"
    End If
End Sub


Public Function GasString(gasValue As Long) As String
' ----------------------------------------------------------------------------
' строковое значение газа
' Last update: 19.04.2018
' ----------------------------------------------------------------------------
    Select Case gasValue
        Case 0
            GasString = "Отсутствует"
        Case 1
            GasString = "Сетевой"
        Case 2
            GasString = "Баллонный"
        Case Else
            GasString = "#Ошибка"
    End Select
End Function


Public Function HeatingString(heatingValue As Long) As String
' ----------------------------------------------------------------------------
' строковое значение отопления
' Last update: 19.04.2018
' ----------------------------------------------------------------------------
    Select Case heatingValue
        Case 0
            HeatingString = "Отсутствует"
        Case 1
            HeatingString = "Центральное"
        Case 2
            HeatingString = "Индивидуальное"
        Case Else
            HeatingString = "#Ошибка"
    End Select
End Function


Public Function HotWaterString(hwValue As Long) As String
' ----------------------------------------------------------------------------
' строковое значение горячего водоснабжения
' Last update: 19.04.2018
' ----------------------------------------------------------------------------
    Select Case hwValue
        Case 0
            HotWaterString = "Отсутствует"
        Case 1
            HotWaterString = "Открытая"
        Case 2
            HotWaterString = "Закрытая"
        Case Else
            HotWaterString = "#Ошибка"
    End Select
End Function


Public Function MonthNames(monthnumber As Long) As String
' ----------------------------------------------------------------------------
' названия месяцев в родительном падеже
' Last update: 23.05.2018
' ----------------------------------------------------------------------------
    Select Case monthnumber
        Case 1:
            MonthNames = "января"
        Case 2:
            MonthNames = "февраля"
        Case 3:
            MonthNames = "марта"
        Case 4:
            MonthNames = "апреля"
        Case 5:
            MonthNames = "мая"
        Case 6:
            MonthNames = "июня"
        Case 7:
            MonthNames = "июля"
        Case 8:
            MonthNames = "августа"
        Case 9:
            MonthNames = "сентября"
        Case 10:
            MonthNames = "октября"
        Case 11:
            MonthNames = "ноября"
        Case 12:
            MonthNames = "декабря"
    End Select
End Function


Public Function CollectionToArray(myCol As Collection) As Variant
' ----------------------------------------------------------------------------
' преобразование коллекции в массив
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    Dim result  As Variant
    Dim Cnt     As Long

    ReDim result(myCol.count - 1)

    For Cnt = 0 To myCol.count - 1
        result(Cnt) = myCol(Cnt + 1)
    Next Cnt

    CollectionToArray = result
End Function


Public Function getFileText(fileName As String) As String
' ----------------------------------------------------------------------------
' Возврат содержимого текстового файла
' Last update: 17.03.2021
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim fso As Object, fsFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsFile = fso.OpenTextFile(fileName)
    If fsFile.AtEndOfStream Then
        getFileText = ""
    Else
        getFileText = fsFile.ReadAll
    End If
    fsFile.Close
    
    Set fsFile = Nothing
    Set fso = Nothing
    
errHandler:
    If Err.Number = 53 Or Err.Number = 76 Then
        Err.Raise ERROR_FILE_NOT_FOUND, "getFileText", "Файл " & _
                                                    fileName & " не найден"
    ElseIf Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function


Public Sub setFileText(fileName As String, newValues As String)
' ----------------------------------------------------------------------------
' сохранение текста newValues в файл fileName
' Last update: 28.11.2018
' ----------------------------------------------------------------------------
    Dim fNumber As Integer
    Dim pathFile As String
    
    On Error GoTo errHandler
    pathFile = Left(fileName, InStrRev(fileName, Application.PathSeparator))
    If Dir(pathFile, vbDirectory) = "" Then MkDir pathFile
    
    fNumber = FreeFile
    Open fileName For Output As #fNumber
    Print #fNumber, newValues
    Close #fNumber
    Exit Sub
    
errHandler:
    If Err.Number = 53 Or Err.Number = 76 Then
        Err.Raise ERROR_FILE_NOT_FOUND, "setFileText", "Файл " & _
                                                    fileName & " не найден"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub


Public Sub setSetting(newValue As Variant, settingType As ConstValuesEnum)
' ----------------------------------------------------------------------------
' установка значение на листе Settings
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    ThisWorkbook.Worksheets(shtSettings).Cells(settingType, 1).Value = _
                                                                    newValue
End Sub


Public Function getUserSetting(settingType As ConstValuesEnum) As Variant
' ----------------------------------------------------------------------------
' получение значения с листа Settings
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    getUserSetting = _
        ThisWorkbook.Worksheets(shtSettings).Cells(settingType, 1).Value
End Function


Public Function EscapeJSON(JsonString As String) As String
' ----------------------------------------------------------------------------
' экранирование символов для JSON
' Last update: 24.09.2020
' ----------------------------------------------------------------------------
    If Len(JsonString) = 0 Then
        EscapeJSON = ""
    Else
        EscapeJSON = Replace(Replace(JsonString, "\", "\\"), """", "\""")
    End If
End Function


Public Function NumberToJSON(NumValue As Variant) As String
' ----------------------------------------------------------------------------
' экранирование чисел для JSON
' Last update: 28.12.2020
' ----------------------------------------------------------------------------
    NumberToJSON = Replace(CStr(NumValue), ",", ".")
End Function


Public Sub NewZip(zipFileName As String)
' ----------------------------------------------------------------------------
' создание пустого zip файла
' Last update: 02.03.2021
' ----------------------------------------------------------------------------
    Open zipFileName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Public Function getSimpleFormValue(valueType As ReloadComboMethods, _
            formTitle As String, _
            Optional ByRef getFormText As String) As Long
' ----------------------------------------------------------------------------
' выбор значения с помощью SimpleSelectionForm
' 09.09.2021
' ----------------------------------------------------------------------------
    getSimpleFormValue = NOTVALUE
    With SimpleSelectionForm

        .Caption = formTitle
        Call reloadComboBox(valueType, .ComboBox1)
        .ComboBox1.ListIndex = 0
        .show
        getSimpleFormValue = .selectedItem
        getFormText = .selectedText
    End With
    Unload SimpleSelectionForm
End Function


Function inary22aServiceInCollection(serviceName As String, _
                                    searchCollection As Collection) As Integer
' ----------------------------------------------------------------------------
' функция для загрузки отчета 22. Ищет индекс услуги в коллекции
' Last update: 22.03.2021
' ----------------------------------------------------------------------------
    Dim i As Integer
    For i = 1 To searchCollection.count
        If searchCollection(i).serviceName = serviceName Then
            inary22aServiceInCollection = i
            Exit Function
        End If
    Next i
    inary22aServiceInCollection = NOTVALUE
End Function


Function stringHasRussianLetter(inString As String) As Boolean
' ----------------------------------------------------------------------------
' проверка на наличие в строке русских букв
' Last update: 22.03.2021
' ----------------------------------------------------------------------------
    Dim strIdx As Integer
    Dim curChrCode As Integer
    For strIdx = 1 To Len(inString)
        curChrCode = Asc(mid(inString, strIdx, 1))
        If (curChrCode = 168) Or (curChrCode = 184) Or _
                ((191 < curChrCode) And (curChrCode < 256)) Then
            stringHasRussianLetter = True
            Exit Function
        End If
    Next strIdx
    stringHasRussianLetter = False
End Function


Sub highlightListItem(curItem As listItem, selectionColor As Long)
' ----------------------------------------------------------------------------
' выделение строки ListView указанным цветом
' Last update: 08.04.2021
' ----------------------------------------------------------------------------
    Dim subItemIdx As Long
    
    curItem.ForeColor = selectionColor
    
    For subItemIdx = 1 To curItem.ListSubItems.count
        curItem.ListSubItems(subItemIdx).ForeColor = selectionColor
    Next subItemIdx
    
End Sub


Function isXmlValidHeader(headerType As XmlHeaderTypeEnum, _
        xmlNode As Object) As Boolean
' ----------------------------------------------------------------------------
' проверка заголовка xml на валидность
' 12.08.2021
' ----------------------------------------------------------------------------
    isXmlValidHeader = True
    On Error GoTo errHandler
    
    If CInt(xmlNode.Attributes.GetNamedItem("version").Value) <> _
            AppConfig.xmlVersion(headerType) Then
        isXmlValidHeader = False
        Exit Function
    End If
    
    If StrComp(xmlNode.Attributes.GetNamedItem("type").Value, _
            AppConfig.xmlType(headerType)) <> 0 Then
        isXmlValidHeader = False
        Exit Function
    End If
    Exit Function
    
errHandler:
    isXmlValidHeader = False
    Err.Clear
End Function


Public Sub PickDateToLabel(ByRef DestLabel As MSForms.Label, _
                        ByRef SavesObject As Object, _
                        Optional formTitle As String = "Выберите дату")
' ----------------------------------------------------------------------------
' Выбор даты в указанный Label с сохранением нужного объекта
' (показывается модальная форма, поэтому текущая форма сбрасывается, и если
' в ней есть какой-то текущий объект, то он обнуляется)
' 26.08.2021
' ----------------------------------------------------------------------------
    With DatePickerForm
        Set .Target = DestLabel
        Set .ParentObject = SavesObject
        .Caption = formTitle
        .show
        Set SavesObject = .ParentObject
    End With
    Unload DatePickerForm
End Sub


Function getTmpFileName() As String
' ----------------------------------------------------------------------------
' получить имя временного файла
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim fso As Object
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    getTmpFileName = getThisPath & fso.GetTempName()
    Set fso = Nothing
End Function


Public Function checkControlFillText(curControl As MSForms.Control) As Boolean
' ----------------------------------------------------------------------------
' проверка заполнения текстового элемента
' 20.10.2021
' ----------------------------------------------------------------------------
    checkControlFillText = Len(Trim(curControl.text)) > 0
End Function


Public Function checkControlFillNumber(curControl As MSForms.Control) As Boolean
' ----------------------------------------------------------------------------
' проверка заполнения текстового элемента
' 20.10.2021
' ----------------------------------------------------------------------------
    checkControlFillNumber = IsNumeric(curControl.text)
End Function


Sub AutoResizeListViewColumnHeader(MyListView As Variant, _
            columnNumber As Integer)
' ----------------------------------------------------------------------------
' установка ширины столбца ListView по ширине заголовка
' 23.06.2022
' ----------------------------------------------------------------------------
    Dim myLabel As Object
    Dim MaxColumnWidth As Double
    Dim i As Integer
    
    ' Create a dynamic label and set it invisible
    
    Set myLabel = MyListView.Parent.Controls.add("Forms.Label.1", "Test Label", True)
 
    With myLabel
        .font.Size = MyListView.font.Size
        .font.Name = MyListView.font.Name
        .WordWrap = False
        .AutoSize = True
        .Visible = False
    End With
 
    myLabel.Caption = MyListView.ColumnHeaders(columnNumber).text
    MaxColumnWidth = myLabel.Width
 
    For i = 1 To MyListView.ListItems.count
        myLabel.Caption = MyListView.ListItems(i).ListSubItems(columnNumber - 1).text
        If myLabel.Width > MaxColumnWidth Then
            MaxColumnWidth = myLabel.Width
        End If
    Next i

    MyListView.ColumnHeaders(columnNumber).Width = MaxColumnWidth + 8
    
    MyListView.Parent.Controls.remove myLabel.Name    'Remove the dynamic label

End Sub


Function FileToByteArray(ByVal sName As String) As Byte()
' ----------------------------------------------------------------------------
' преобразование файла в массив байт
' 02.10.2022
' ----------------------------------------------------------------------------
    Dim b() As Byte
    Dim n As Long
    n = FileLen(sName)
    
    ReDim b(n - 1)
    
    Open sName For Binary Access Read As #1
    Get #1, , b()
    Close #1
    FileToByteArray = b
End Function


Function FileFromByteArray(bArray() As Byte) As String
' ----------------------------------------------------------------------------
' преобразование массив байт во временный файл и возврат этого файла
' 02.10.2022
' ----------------------------------------------------------------------------
    Dim n As String
    n = getTmpFileName
        
    Open n For Binary Access Write As #1
    Put #1, , bArray
    Close #1
    FileFromByteArray = n
End Function
