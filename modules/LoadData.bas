Attribute VB_Name = "LoadData"
Option Explicit
Option Private Module

Enum ElectroFileColumns
    elBldnId = 4
    elDate = 6
    elFlat = 11
    elReadings = 25
End Enum

Enum HotWaterFileColumns
    hwBldnId = 10
    hwFlat = 5
    hwVolume = 8
    hwNormVolume = 12
    hwMeterVolume = 13
End Enum


Enum FsoIoMode
    ForReading = 1      ' только чтение
    ForWriting = 2      ' запись с очисткой
    ForAppending = 8    ' добавление
End Enum


Enum FsoFormat
    TristateUseDefault = -2 ' системные настройки
    TristateTrue = -1       ' UTF
    TristateFalse = 0       ' ASCII
End Enum

Private Enum rkcReport8Columns
    rrc_Service = 1
    rrc_Contractor
    rrc_MD
    rrc_Village
    rrc_Address
    rrc_OccId
    rrc_Flat
    rrc_FIO
    rrc_Type
    rrc_Document
    rrc_Sum
    rrc_Date
    rrc_Volume
End Enum


Sub loadMeterReadings()
' ----------------------------------------------------------------------------
' загрузка показаний ИПУ
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    Dim screenUpdStatus As Boolean
    screenUpdStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    Dim errMsg As String
    
    ' Выбор услуги
    Dim serviceId As Long
    Dim serviceName As String
    serviceId = getSimpleFormValue(rcmServices, "Выберите услугу")
    serviceName = LCase(services.ServiceNameById(serviceId))
    
    ' выбор периода для ГВС
    Dim curTermId As Long
    If services.IsHotWater(serviceId) Then
        curTermId = getSimpleFormValue(rcmTermDESC, _
                "Укажите месяц для показаний " & serviceName)
    End If
    
    ' открытие книги
    Dim wbName As String
    Dim wb As Workbook
    Dim strTitle As String
    strTitle = "Выберите файл с показаниями " & serviceName
    If services.IsHotWater(serviceId) Then
        strTitle = strTitle & " за " & LCase(Terms(CStr(curTermId)).StringValue)
    End If
    wbName = Application.GetOpenFilename( _
                    "Excel файл отчёта (*.xls;*.xlsx),*.xls;*.xlsx", _
                    Title:=strTitle)
    If wbName = "False" Then
        errMsg = "Операция отменена"
        GoTo errHandler
    End If
    Set wb = Workbooks.Open(wbName, ReadOnly:=True)
    
    ' создание текстового файла, куда будет записываться информация
    Dim fso As Object, tmpFile As Object
    Dim fileName As String
    Dim thisPCFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetTempName()
    thisPCFileName = getThisPath & fileName
    Set tmpFile = fso.CreateTextFile(thisPCFileName)
    
    ' заполнение файла
    If services.IsElectro(serviceId) Then
        Call createElectroLoadFile(wb, tmpFile, serviceId)
    ElseIf services.IsHotWater(serviceId) Then
        Call createHotWaterLoadFile(wb, tmpFile, serviceId, curTermId)
    End If
    
    tmpFile.Close
    
    FileCopy thisPCFileName, AppConfig.ServerTmpPath & fileName
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    Dim conn As New DBAdmConnection
    Dim transId As Long
    
    sqlString = "load_meter_readings"
    sqlParams.add "InFileName", fileName
    conn.initial DB_NAME
    transId = conn.Connection.BeginTrans
    conn.RunQuery sqlString, sqlParams
    conn.Connection.CommitTrans
    MsgBox "Успешно завершено"
    GoTo cleanHandler

errHandler:
    If errMsg = "" And Err.Number <> 0 Then
        errMsg = Err.Description
    End If
    If transId > 0 Then conn.Connection.RollbackTrans
    MsgBox errMsg, vbOKOnly
    
cleanHandler:
    If Not conn Is Nothing Then conn.Connection.Close
    Set conn = Nothing
    Set sqlParams = Nothing
    If Not wb Is Nothing Then wb.Close savechanges:=False
    Set wb = Nothing
    If Not tmpFile Is Nothing Then tmpFile.Close
    If fileName <> "" Then
        If fso.FileExists(thisPCFileName) Then fso.DeleteFile (thisPCFileName)
        If fso.FileExists(AppConfig.ServerTmpPath & fileName) Then fso.DeleteFile (AppConfig.ServerTmpPath & fileName)
    End If
    Set tmpFile = Nothing
    Set fso = Nothing
    
    Application.ScreenUpdating = screenUpdStatus
End Sub


Sub createHotWaterLoadFile(ByRef InWorkBook As Workbook, _
                            ByRef OutTmpFile As Object, _
                            InServiceId As Long, _
                            InTermId As Long)
' ----------------------------------------------------------------------------
' создание файла с показаниями ИПУ ГВС
' Last update: 30.07.2021
' ----------------------------------------------------------------------------
    With InWorkBook.ActiveSheet
        Dim rowIdx As Long
        Dim curBldnId As Long
        Dim curFlat As String
        
        For rowIdx = 1 To .UsedRange.Rows.count
            ' выбираются только строки, где код дома число, для пропуска пустых
            curBldnId = longValue(.Cells(rowIdx, HotWaterFileColumns.hwBldnId).Value)
            curFlat = .Cells(rowIdx, HotWaterFileColumns.hwFlat).Value
            If curBldnId > 0 And curFlat <> "" Then
                OutTmpFile.writeline _
                    curBldnId & _
                    ";" & curFlat & _
                    ";" & InTermId & _
                    ";" & NumberToJSON(dblValue(.Cells(rowIdx, HotWaterFileColumns.hwMeterVolume).Value) + _
                            dblValue(.Cells(rowIdx, HotWaterFileColumns.hwNormVolume).Value) + _
                            dblValue(.Cells(rowIdx, HotWaterFileColumns.hwVolume).Value)) & _
                    ";" & InServiceId
            End If      ' if curbldn > 0
        Next rowIdx
    End With
End Sub


Sub createElectroLoadFile(ByRef InWorkBook As Workbook, _
                            ByRef OutTmpFile As Object, _
                            InServiceId As Long)
' ----------------------------------------------------------------------------
' создание файла с показаниями ИПУ ЭЭ
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    Dim errElectroId As New MyCollection
    
    With InWorkBook.ActiveSheet
        Dim rowIdx As Long
        Dim curBldnId As Long, curOurBldnId As Long
        Dim mappingBldn As New bldn_mapping
        
        For rowIdx = 1 To .UsedRange.Rows.count
            ' выбираются только строки, где код дома число, для пропуска пустых
            curBldnId = longValue(.Cells(rowIdx, ElectroFileColumns.elBldnId).Value)
            If curBldnId > 0 Then
                curOurBldnId = mappingBldn.GetOurIdByElectro(curBldnId)
                If curOurBldnId = NOTVALUE Then
                    If Not errElectroId.elementInCollection(curBldnId) Then
                        errElectroId.add curBldnId
                        MsgBox "Дом из файла с кодом " & curBldnId & _
                            " не сопоставлен с домом в базе", vbExclamation, "Внимание"
                    End If  ' curbldnid in errCollection
                Else
                    OutTmpFile.writeline curOurBldnId & _
                        ";" & .Cells(rowIdx, ElectroFileColumns.elFlat).Value & _
                        ";" & Terms.TermByDate(.Cells(rowIdx, ElectroFileColumns.elDate).Value).Id & _
                        ";" & .Cells(rowIdx, ElectroFileColumns.elReadings).Value & _
                        ";" & InServiceId
                End If      ' curOurBldnId = notvalue
            End If      ' if curbldn > 0
        Next rowIdx
    End With
    Set errElectroId = Nothing
End Sub


Sub loadAccruedToBase(infoType As AccruedTypes)
' ----------------------------------------------------------------------------
' Загрузка начислений по домам
' Last update: 25.03.2021
' ----------------------------------------------------------------------------
    Dim fileLoadName As String
    Dim errMsg As String
    
    
    ' период, за который загружается информация
    Dim curTermId As Long
    curTermId = getSimpleFormValue(rcmTermDESC, "Выберите период")
    
    ' создание файла для загрузки и получение его имени
    fileLoadName = createAccruedFile(infoType, curTermId, errMsg)
    
    If fileLoadName = NOTSTRING Then
        MsgBox errMsg, vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' загрузка информации в базу
    On Error Resume Next
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    Dim conn As New DBAdmConnection
    
    sqlString = "load_rkc_values"
    sqlParams.add "InFileName", fileLoadName
    sqlParams.add "InTermId", curTermId
    sqlParams.add "InSourceType", infoType
    conn.initial DB_NAME
    conn.Connection.CommandTimeout = 0
    conn.Connection.BeginTrans
    conn.RunQuery sqlString, sqlParams
    If Err.Number <> 0 Then
        conn.Connection.RollbackTrans
        MsgBox Err.Description, vbExclamation, "Ошибка"
    Else
        conn.Connection.CommitTrans
        MsgBox "Успешно завершено"
    End If
    conn.Connection.Close
    If Dir(getThisPath & fileLoadName) <> "" Then Kill getThisPath & fileLoadName
    If Dir(AppConfig.ServerTmpPath & fileLoadName) <> "" Then Kill AppConfig.ServerTmpPath & fileLoadName
    Set conn = Nothing
    Set sqlParams = Nothing

End Sub


Function createAccruedFile(accruedType As AccruedTypes, _
                            curTermId As Long, _
                            ByRef errMsg As String) As String
' ----------------------------------------------------------------------------
' создание файла с начислениями
' Last update: 25.03.2021
' ----------------------------------------------------------------------------
        
    Dim screenUpdStatus As Boolean
    screenUpdStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    createAccruedFile = NOTSTRING
    
    On Error GoTo errHandler
    
    ' создание текстового файла, куда будет записываться информация по домам
    ' потом он же будет загружаться в базу
    Dim fso As Object, tmpFile As Object
    Dim fileName As String
    Dim thisPCFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetTempName()
    thisPCFileName = getThisPath & fileName
    Set tmpFile = fso.CreateTextFile(thisPCFileName)
    
    Select Case accruedType
        Case AccruedTypes.acMOKR
            Call fillMOKapRemont(tmpFile, curTermId, errMsg)
        Case AccruedTypes.acBuh
            Call fillYurLico(tmpFile, curTermId, errMsg)
        Case AccruedTypes.acRKC
            Call fillFromRkc(tmpFile, curTermId, errMsg)
        Case AccruedTypes.acMOUpper
            Call fillMOUpper(tmpFile, curTermId, errMsg)
        Case Else
            errMsg = "Неверный выбор типа файла"
    End Select
    If errMsg <> "" Then GoTo errHandler
    
    tmpFile.Close
    
    FileCopy thisPCFileName, AppConfig.ServerTmpPath & fileName
     
    createAccruedFile = fileName
    GoTo cleanHandler

errHandler:
    If errMsg = "" And Err.Number <> 0 Then
        errMsg = Err.Description
    End If
    If Not tmpFile Is Nothing Then tmpFile.Close
    If fso.FileExists(thisPCFileName) Then fso.DeleteFile (thisPCFileName)
    If fso.FileExists(AppConfig.ServerTmpPath & fileName) Then fso.DeleteFile (AppConfig.ServerTmpPath & fileName)
    
cleanHandler:
    
    Set tmpFile = Nothing
    Set fso = Nothing
    
    Application.ScreenUpdating = screenUpdStatus

End Function


Sub fillFromRkc(ByRef tmpFile As Object, _
                        termId As Long, _
                        ByRef errMsg As String)
' ----------------------------------------------------------------------------
' заполнение файла с начислениями РКЦ
' Last update: 19.01.2022
' ----------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Dim accruedType As AccruedTypes
    accruedType = acRKC
    
    ' открытие книги отчета 22а
    Dim wb22aName As String
    Dim wb22a As Workbook
    wb22aName = Application.GetOpenFilename( _
                    "Excel файл отчёта (*.xls;*.xlsx),*.xls;*.xlsx", _
                    Title:="Выберите файл отчета 22а за " & _
                                        LCase(Terms(CStr(termId)).StringValue))
    If wb22aName = "False" Then
        errMsg = "Операция отменена"
        GoTo errHandler
    End If
    Set wb22a = Workbooks.Open(wb22aName, ReadOnly:=True)
    
    ' услуги РКЦ, чтобы записывать в файл их код
    Dim rkcServices As New rkc_services
    rkcServices.reload
    
    With wb22a.ActiveSheet
        Dim bService As Inary22aInfo
        Dim rowIdx As Long
        Dim curBldnId As Long
        Dim serviceName As String

        For rowIdx = 1 To .UsedRange.Rows.count
            ' выбираются только строки, где код дома число, для пропуска пустых
            curBldnId = longValue(.Cells(rowIdx, Inary22aColumns.i22aBldnId).Value)
            If curBldnId > 0 Then
                serviceName = .Cells(rowIdx, Inary22aColumns.i22aService).Value
                If StrComp(serviceName, "Итого", vbTextCompare) <> 0 Then
                    If rkcServices.GetServiceIdByName(serviceName) = NOTVALUE Then
                        errMsg = "Услуга " & serviceName & " не найдена в базе"
                        GoTo errHandler
                    End If
                    Set bService = New Inary22aInfo
                    bService.BldnId = curBldnId
                    bService.serviceName = serviceName
                    bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                    bService.termId = termId
                    bService.accruedType = accruedType
                    bService.Accrued = .Cells(rowIdx, Inary22aColumns.i22aAccrued).Value
                    bService.Added = .Cells(rowIdx, Inary22aColumns.i22aAdded).Value
                    bService.Compens = .Cells(rowIdx, Inary22aColumns.i22aCompens).Value
                    bService.Paid = .Cells(rowIdx, Inary22aColumns.i22aPaid).Value
                    bService.flatNo = .Cells(rowIdx, Inary22aColumns.i22aFlatNo).Value
                    bService.OccId = .Cells(rowIdx, Inary22aColumns.i22aOccId).Value
                    bService.InSaldo = .Cells(rowIdx, Inary22aColumns.i22aInSaldo).Value
                    bService.OutSaldo = .Cells(rowIdx, Inary22aColumns.i22aOutSaldo).Value
                    If bService.haveData Then
                        tmpFile.writeline bService.ExportString
                    End If      ' if haveData
                End If      ' if <> "Итого"
            End If      ' if curbldn > 0
        Next rowIdx
    End With
    
    GoTo cleanHandler

errHandler:
    If errMsg = "" And Err.Number <> 0 Then
        errMsg = Err.Description
    End If
    
cleanHandler:
    If Not wb22a Is Nothing Then wb22a.Close savechanges:=False
    Set wb22a = Nothing
    Set bService = Nothing
    Set rkcServices = Nothing
End Sub


Sub fillYurLico(ByRef tmpFile As Object, _
                    termId As Long, _
                    ByRef errMsg As String)
' ----------------------------------------------------------------------------
' заполнение файла с начислениями по офисам
' Last update: 25.01.2022
' ----------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Const SODERZHANIE_NAME$ = "Сод.жил.по"
    Const SOI_HOT_WATER$ = "ГВ СОИ"
    Const SOI_COLD_WATER$ = "ХВ СОИ"
    Const SOI_ELECTRO$ = "Эл.эн. СОИ"
    
    Dim accruedType As AccruedTypes
    accruedType = acBuh
    
    ' открытие книги
    Dim wb22aName As String
    Dim wb22a As Workbook
    wb22aName = Application.GetOpenFilename( _
                    "Excel файл отчёта (*.xls;*.xlsx),*.xls;*.xlsx", _
                    Title:="Выберите файл с начислениями офисам за " & _
                            LCase(Terms(CStr(termId)).StringValue))
    If wb22aName = "False" Then
        errMsg = "Операция отменена"
        GoTo errHandler
    End If
    Set wb22a = Workbooks.Open(wb22aName, ReadOnly:=True)
    
    ' услуги РКЦ, чтобы записывать в файл их код
    Dim rkcServices As New rkc_services
    rkcServices.reload
    
    With wb22a.ActiveSheet
        Dim bServices As Collection
        Dim bService As Inary22aInfo
        Dim rowIdx As Long
        Dim curBldnId As Long
        Dim serviceId As Long
        
        For rowIdx = 1 To .UsedRange.Rows.count
            ' выбираются только строки, где код дома число, для пропуска пустых
            curBldnId = longValue(.Cells(rowIdx, YurLicoSheetColumns.ylscBldnId).Value)
            If curBldnId > 0 Then
                ' все начисления на одной строке, поэтому загружаем все услуги
                Set bServices = New Collection
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SODERZHANIE_NAME
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.flatNo = .Cells(rowIdx, YurLicoSheetColumns.ylscFlatNo).Value
                bService.OccId = .Cells(rowIdx, YurLicoSheetColumns.ylscOccId).Value
                bService.Accrued = .Cells(rowIdx, YurLicoSheetColumns.ylscSodAccrued).Value
                bService.Paid = .Cells(rowIdx, YurLicoSheetColumns.ylscSodPaid).Value
                bServices.add bService
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SOI_ELECTRO
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.flatNo = .Cells(rowIdx, YurLicoSheetColumns.ylscFlatNo).Value
                bService.OccId = .Cells(rowIdx, YurLicoSheetColumns.ylscOccId).Value
                bService.Accrued = .Cells(rowIdx, YurLicoSheetColumns.ylscElectroAccrued).Value
                bService.Paid = .Cells(rowIdx, YurLicoSheetColumns.ylscElectroPaid).Value
                bServices.add bService
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SOI_COLD_WATER
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.flatNo = .Cells(rowIdx, YurLicoSheetColumns.ylscFlatNo).Value
                bService.OccId = .Cells(rowIdx, YurLicoSheetColumns.ylscOccId).Value
                bService.Accrued = .Cells(rowIdx, YurLicoSheetColumns.ylscColdWaterAccrued).Value
                bService.Paid = .Cells(rowIdx, YurLicoSheetColumns.ylscColdWaterPaid).Value
                bServices.add bService
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SOI_HOT_WATER
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.flatNo = .Cells(rowIdx, YurLicoSheetColumns.ylscFlatNo).Value
                bService.OccId = .Cells(rowIdx, YurLicoSheetColumns.ylscOccId).Value
                bService.Accrued = .Cells(rowIdx, YurLicoSheetColumns.ylscHotWaterAccrued).Value
                bService.Paid = .Cells(rowIdx, YurLicoSheetColumns.ylscHotWaterPaid).Value
                bServices.add bService
                
                For Each bService In bServices
                    If bService.haveData Then
                        tmpFile.writeline bService.ExportString
                    End If      ' if haveData
                Next bService
            End If      ' if curbldn > 0
        Next rowIdx
        tmpFile.Close
    End With
    
    GoTo cleanHandler

errHandler:
    If errMsg = "" And Err.Number <> 0 Then
        errMsg = Err.Description
    End If
    
cleanHandler:
    If Not wb22a Is Nothing Then wb22a.Close savechanges:=False
    Set wb22a = Nothing
    Set bService = Nothing
    Set rkcServices = Nothing
End Sub


Sub fillMOKapRemont(tmpFile As Object, _
                        termId As Long, _
                        ByRef errMsg As String)
' ----------------------------------------------------------------------------
' заполнение файла с начислениями взноса на капремонт
' 13.04.2022
' ----------------------------------------------------------------------------
    
    Const SODERZHANIE_NAME$ = "Вз.кап.рем"
    
    On Error GoTo errHandler
    
    Dim accruedType As AccruedTypes
    accruedType = acMOKR
    
    ' открытие книги
    Dim wb22aName As String
    Dim wb22a As Workbook
    wb22aName = Application.GetOpenFilename( _
                    "Excel файл отчёта (*.xls;*.xlsx),*.xls;*.xlsx", _
                    Title:="Выберите файл начислений взноса на КР для МО за " & _
                            LCase(Terms(CStr(termId)).StringValue))
    If wb22aName = "False" Then
        errMsg = "Операция отменена"
        GoTo errHandler
    End If
    Set wb22a = Workbooks.Open(wb22aName, ReadOnly:=True)
    
    ' услуги РКЦ, чтобы записывать в файл их код
    Dim rkcServices As New rkc_services
    rkcServices.reload
    
    With wb22a.ActiveSheet
        Dim bService As Inary22aInfo
        Dim rowIdx As Long
        Dim curBldnId As Long
        Dim serviceId As Long
        
        For rowIdx = 1 To .UsedRange.Rows.count
            ' выбираются только строки, где код дома число, для пропуска пустых
            curBldnId = longValue(.Cells(rowIdx, MOKarRemontSheetColumns.mkrBldnId).Value)
            If curBldnId > 0 Then
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SODERZHANIE_NAME
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.OccId = .Cells(rowIdx, MOKarRemontSheetColumns.mkrOccId).Value
                bService.Accrued = .Cells(rowIdx, MOKarRemontSheetColumns.mkrAccrued).Value
                bService.Paid = .Cells(rowIdx, MOKarRemontSheetColumns.mkrPaid).Value
                bService.flatNo = .Cells(rowIdx, MOKarRemontSheetColumns.mkrFlatNo).Value
                    
                If bService.haveData Then
                    tmpFile.writeline bService.ExportString
                End If      ' if haveData
            End If      ' if curbldn > 0
        Next rowIdx
        tmpFile.Close
    End With

errHandler:
    errMsg = Err.Description
    
cleanHandler:
    If Not wb22a Is Nothing Then wb22a.Close savechanges:=False
    Set wb22a = Nothing
    Set bService = Nothing
    Set rkcServices = Nothing
End Sub


Sub fillMOUpper(ByRef tmpFile As Object, _
                    termId As Long, _
                    ByRef errMsg As String)
' ----------------------------------------------------------------------------
' заполнение файла с начислениями превышение по офисам
' Last update: 25.03.2021
' ----------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Const SODERZHANIE_NAME$ = "Сод.жил.по"
    Const SOI_HOT_WATER$ = "ГВ СОИ"
    Const SOI_COLD_WATER$ = "ХВ СОИ"
    
    Dim accruedType As AccruedTypes
    accruedType = acMOUpper
    
    ' открытие книги
    Dim wb22aName As String
    Dim wb22a As Workbook
    wb22aName = Application.GetOpenFilename( _
                    "Excel файл отчёта (*.xls;*.xlsx),*.xls;*.xlsx", _
                    Title:="Выберите файл превышений МО за " & _
                            LCase(Terms(CStr(termId)).StringValue))
    If wb22aName = "False" Then
        errMsg = "Операция отменена"
        GoTo errHandler
    End If
    Set wb22a = Workbooks.Open(wb22aName, ReadOnly:=True)
    
    ' услуги РКЦ, чтобы записывать в файл их код
    Dim rkcServices As New rkc_services
    rkcServices.reload
    
    With wb22a.ActiveSheet
        Dim bServices As Collection
        Dim bService As Inary22aInfo
        Dim rowIdx As Long
        Dim curBldnId As Long
        Dim serviceId As Long
        
        For rowIdx = 1 To .UsedRange.Rows.count
            ' выбираются только строки, где код дома число, для пропуска пустых
            curBldnId = longValue(.Cells(rowIdx, MoUpperSheetColumns.muBldnId).Value)
            If curBldnId > 0 Then
                ' все начисления на одной строке, поэтому загружаем все услуги
                Set bServices = New Collection
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SODERZHANIE_NAME
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.Accrued = .Cells(rowIdx, MoUpperSheetColumns.muSodAccured).Value
                bService.Paid = .Cells(rowIdx, MoUpperSheetColumns.muSodPaid).Value
                bServices.add bService
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SOI_COLD_WATER
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.Accrued = .Cells(rowIdx, MoUpperSheetColumns.muColdWaterAccrued).Value
                bService.Paid = .Cells(rowIdx, MoUpperSheetColumns.muColdWaterPaid).Value
                bServices.add bService
                
                Set bService = New Inary22aInfo
                bService.BldnId = curBldnId
                bService.termId = termId
                bService.accruedType = accruedType
                bService.serviceName = SOI_HOT_WATER
                bService.serviceId = rkcServices.GetServiceIdByName(bService.serviceName)
                bService.Accrued = .Cells(rowIdx, MoUpperSheetColumns.muHotWaterAccrued).Value
                bService.Paid = .Cells(rowIdx, MoUpperSheetColumns.muHotWaterPaid).Value
                bServices.add bService
                
                For Each bService In bServices
                    If bService.haveData Then
                        tmpFile.writeline bService.ExportString
                    End If      ' if haveData
                Next bService
            End If      ' if curbldn > 0
        Next rowIdx
        tmpFile.Close
    End With
    
    GoTo cleanHandler

errHandler:
    If errMsg = "" And Err.Number <> 0 Then
        errMsg = Err.Description
    End If
    
cleanHandler:
    If Not wb22a Is Nothing Then wb22a.Close savechanges:=False
    Set wb22a = Nothing
    Set bService = Nothing
    Set rkcServices = Nothing
End Sub


Sub loadRkcAddeds()
' ----------------------------------------------------------------------------
' загрузка разовых начислений
' 09.09.2021
' ----------------------------------------------------------------------------
    Dim loadFileName As String, xmlFileName As String
    Dim curTermId As Long, curTypeId As Long
    Dim titulString As String, outFormString As String
    Dim errMsg As String
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    titulString = "Выберите файл "
    curTermId = getSimpleFormValue(rcmTermDESC, "Выберите период", _
                                outFormString)
    titulString = titulString & outFormString
    curTypeId = getSimpleFormValue(rcmAddedTypes, "Выберите тип разовых", _
                                outFormString)
    titulString = titulString & " " & outFormString
        
    loadFileName = Application.GetOpenFilename( _
                "Excel-файлы (*.xls; *.xlsx),*.xls;*.xlsx", , titulString)
    If StrComp(loadFileName, "False", vbTextCompare) = 0 Then
        MsgBox "Операция отменена"
        GoTo cleanHandler
    End If
    
    On Error GoTo errHandler
    
    ' создание xml-файла для загрузки
    Dim loadWB As Workbook, loadWS As Worksheet
    Dim xml As Object, root As Object, curItem As Object
    Dim xml_file_name As String
    Dim i As Long
    Dim serviceName As String, serviceId As Long, curOcc As Long
    Dim rkcServices As New rkc_services
    
    Set loadWB = Workbooks.Open(loadFileName, ReadOnly:=True)
    Set loadWS = loadWB.Sheets(1)
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    xml.AppendChild xml.createProcessingInstruction("xml", "version='1.0' encoding='utf-8'")
    Set root = xml.createElement("addeds")
    root.SetAttribute "version", "1"
    root.SetAttribute "term", curTermId
    root.SetAttribute "type", curTypeId
    xml.AppendChild root
    
    With loadWS
        For i = 2 To .UsedRange.Rows.count
            curOcc = longValue(.Cells(i, rkcReport8Columns.rrc_OccId).Value)
            If curOcc <> 0 Then
                If .Cells(i, rkcReport8Columns.rrc_Document).Value = "всего по лс" Then
                    serviceName = .Cells(i, rkcReport8Columns.rrc_Service).Value
                    serviceId = rkcServices.GetServiceIdByFullName(serviceName)
                    If serviceId = NOTVALUE Then
                        errMsg = "Услуга " & serviceName & " не найдена в базе"
                        GoTo errHandler
                    End If  ' getServiceIdByFullName = NOTVALUE
                    
                    Set curItem = xml.createElement("added")
                    curItem.AppendChild(xml.createElement("occ_id")).text = curOcc
                    curItem.AppendChild(xml.createElement("sum")).text = .Cells(i, rkcReport8Columns.rrc_Sum).Value
                    curItem.AppendChild(xml.createElement("service_name")).text = serviceName
                    curItem.AppendChild(xml.createElement("service_id")).text = serviceId
                    root.AppendChild curItem
                End If      ' всего по лс
            End If          ' curocc <> notvalue
        Next i
    End With

    xml_file_name = getTmpFileName()
    xml.Save xml_file_name
    
    ' загрузка файла в базу
    xmlFileName = CopyFileToServer(xml_file_name)
    
    On Error Resume Next
    Dim conn As New DBAdmConnection
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    sqlString = "load_rkc_addeds"
    sqlParams.add "InFileName", xmlFileName
    conn.initial DB_NAME
    conn.Connection.BeginTrans
    conn.RunQuery sqlString, sqlParams
    If Err.Number <> 0 Then
        conn.Connection.RollbackTrans
        GoTo errHandler
    Else
        conn.Connection.CommitTrans
        MsgBox "Успешно завершено"
    End If
    
errHandler:
    If Err.Number <> 0 Then errMsg = Err.Description
    If errMsg <> "" Then MsgBox errMsg, vbExclamation, "Ошибка"
    
cleanHandler:
    If Not conn Is Nothing Then
        If conn.Connection.State = adStateOpen Then
            conn.Connection.Close
        End If
    End If
    If xmlFileName <> "" Then Kill xmlFileName
    If xml_file_name <> "" Then Kill xml_file_name
    If Not loadWB Is Nothing Then loadWB.Close savechanges:=False
        
    Set sqlParams = Nothing
    Set conn = Nothing
    Set xml = Nothing
    Set loadWB = Nothing
    Set loadWS = Nothing
    Set root = Nothing
    Set curItem = Nothing
    Set rkcServices = Nothing
    
    Application.ScreenUpdating = SUStatus
End Sub


Sub loadOffersWorks()
' ----------------------------------------------------------------------------
' загрузка предложений по работам
' 15.10.2021
' ----------------------------------------------------------------------------
    Dim loadFileName As String
    
    loadFileName = Application.GetOpenFilename( _
                "xml-файлы (*.xml),*.xml", , _
                "Выберите файл с информацией о квартирах")
    
    If StrComp(loadFileName, "False", vbTextCompare) = 0 Then
        MsgBox "Операция отменена"
        Exit Sub
    End If
    
    ' проверка на валидность файла
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    If Not xml.Load(loadFileName) Then
        MsgBox "Файл " & loadFileName & " не является xml-файлом", _
                vbExclamation, "Ошибка обработки файла"
        GoTo cleanHandler
    End If
    
    Dim fileNameToLoad As String
    Dim serverFileName As String
    fileNameToLoad = CopyFileToServer(loadFileName)
    serverFileName = AppConfig.ServerTmpPath & Application.PathSeparator & loadFileName
    
    On Error Resume Next
    
    Dim conn As New DBAdmConnection
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    sqlString = "load_offers_works"
    sqlParams.add "InFileName", fileNameToLoad
    conn.initial DB_NAME
    conn.Connection.BeginTrans
    conn.Connection.CommandTimeout = 300
    conn.RunQuery sqlString, sqlParams
    If Err.Number <> 0 Then
        conn.Connection.RollbackTrans
        MsgBox Err.Description, vbExclamation, "Ошибка"
    Else
        conn.Connection.CommitTrans
        MsgBox "Успешно завершено"
    End If
    conn.Connection.Close
    If Dir(serverFileName) <> "" Then Kill serverFileName
    Set sqlParams = Nothing
    
cleanHandler:
    Set xml = Nothing
    
End Sub

Function createTmpFile() As Object
' ----------------------------------------------------------------------------
' создание временного файла (возвращает FileObject)
' 03.08.2021
' ----------------------------------------------------------------------------
    Dim fso As Object
    Dim fileName As String, thisPCFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    thisPCFileName = getTmpFileName()
    fso.CreateTextFile thisPCFileName
    Set createTmpFile = fso.GetFile(thisPCFileName)
    Set fso = Nothing
End Function


Function CopyFileToServer(fileName As String) As String
' ----------------------------------------------------------------------------
' копирование файла fileName в папку на сервере с созданием ему временного
' имени
' 09.08.2021
' ----------------------------------------------------------------------------
    Dim fso As Object
    Dim tmpFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmpFileName = fso.GetTempName()
    
    fso.CopyFile fileName, AppConfig.ServerTmpPath & tmpFileName
    
    CopyFileToServer = AppConfig.InServerTmpPath & tmpFileName
    
    Set fso = Nothing
End Function
