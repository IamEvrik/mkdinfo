Attribute VB_Name = "ModuleFunctions"
Option Explicit
Option Private Module

' ----------------------------------------------------------------------------
' Name: unprotectVBE method
' Parameters: wb - workbook - книга
'             password - string - пароль
' Last update: 13.09.2016
' About: сн€тие парол€ с проекта в книге
' ----------------------------------------------------------------------------
Private Sub unprotectVBE(wb As Workbook, password As String)
    Dim objWindow As VBIDE.Window
    
    For Each objWindow In wb.VBProject.VBE.Windows
        If objWindow.Type = vbext_wt_ProjectWindow Then
'            DoEvents
            objWindow.visible = True
            objWindow.SetFocus
            Exit For
        End If
    Next objWindow
    SendKeys "~" & password & "~", True
    SendKeys "{ENTER}", True
    Set objWindow = Nothing
    wb.VBProject.VBE.MainWindow.visible = False
End Sub

' ----------------------------------------------------------------------------
' Name: ExportModules method
' Last update: 04.10.2016
' About: выгрузка всех модулей проекта
' ----------------------------------------------------------------------------
Public Sub ExportModules()
    Dim s As VBComponent
    Dim ext As String
    Dim saveDir As String
    
    On Error GoTo errHandler
    
    ' путь сохранени€ модулей
    saveDir = modulesDir()
    
    For Each s In ThisWorkbook.VBProject.VBComponents
        If s.Type = vbext_ct_ClassModule Or s.Type = vbext_ct_MSForm _
                    Or s.Type = vbext_ct_StdModule Or s.Name = "Ёта нига" Then
            Select Case s.Type
                Case vbext_ct_StdModule
                    ext = ".bas"
                Case vbext_ct_ClassModule
                    ext = ".cls"
                Case vbext_ct_MSForm
                    ext = ".frm"
            End Select
            s.Export saveDir & s.Name & ext
        End If
    Next s
    
    ' сохранение титульного листа
    ThisWorkbook.Worksheets(shtnmTitul).visible = xlSheetVisible
    ThisWorkbook.Worksheets(shtnmTitul).Copy
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs saveDir & shtnmTitul & ".xlsx"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ActiveWorkbook.Close
    ThisWorkbook.Worksheets(shtnmTitul).visible = xlSheetHidden
    
    MsgBox "Ёкспорт успешно завершЄн", vbOKOnly, "”спешное завершение"
    GoTo cleanHandler
    
errHandler:
    MsgBox "Ёкспорт завершЄн с ошибкой:" & vbCrLf & Err.Description, _
                                            vbExclamation, "ќшибка экспорта"
    GoTo cleanHandler

cleanHandler:
    
End Sub

' ----------------------------------------------------------------------------
' Name: UpdateSoft method
' Last update: 13.09.2016
' About: обновление всех модулей
' ----------------------------------------------------------------------------
Public Sub UpdateSoft()
    On Error GoTo errHandler
    
    Dim newFileName As String, tmpFileName As String    ' имена файлов
    Dim newWB As Workbook                               ' файл дл€ обновлени€
    Dim curModule As VBComponent                        '
    Dim tmpModule As String
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    newFileName = ThisWorkbook.FullName
    tmpFileName = ThisWorkbook.Path & Application.PathSeparator & _
                "ќбновление" & "_" & Year(Date) & "_" & Month(Date) & "_" & _
                Day(Date) & "_" & Hour(Time) & "_" & Minute(Time) & ".xlsm"
    ' переименование текущего файла
    ThisWorkbook.SaveAs tmpFileName
    ' создание и открытие копии текущего файла, в котором обнов€тс€ модули
    ThisWorkbook.SaveCopyAs newFileName
    Set newWB = Workbooks.Open(newFileName)
    ' удаление старых и запись новых модулей в новую книгу
    Call unprotectVBE(newWB, VBA_PASS)
    For Each curModule In newWB.VBProject.VBComponents
        If curModule.Type = vbext_ct_ClassModule Or _
                                curModule.Type = vbext_ct_MSForm _
                                Or curModule.Type = vbext_ct_StdModule Then
            newWB.VBProject.VBComponents.remove curModule
        End If
    Next curModule
    
    tmpModule = Dir(modulesDir())
    While tmpModule <> ""
        If Right(tmpModule, 3) = "frm" Or Right(tmpModule, 3) = "cls" Or _
                Right(tmpModule, 3) = "bas" Or tmpModule = "Ёта нига" Then
            newWB.VBProject.VBComponents.Import modulesDir & tmpModule
        End If
        tmpModule = Dir
    Wend
    
    ' выполнение обновлений из модул€ "updates"
    
    newWB.Close savechanges:=True
    Set newWB = Nothing
    Application.ScreenUpdating = SUStatus
    Exit Sub
    
    ' ќбработка ошибок
    ' удаление временных файлов и возврат текущего
errHandler:
    If Not newWB Is Nothing Then
        newWB.Close savechanges:=False
        Kill newFileName
        Set newWB = Nothing
    End If
    If ThisWorkbook.Name <> newFileName Then
        ThisWorkbook.SaveAs newFileName
        Kill tmpFileName
    End If
    MsgBox "ќшибка обновлени€:" & vbCr & Err.Description
    Err.Clear
    Application.ScreenUpdating = SUStatus
    On Error GoTo 0
        
End Sub

Public Sub DeleteCls(isReally As Boolean)
' ----------------------------------------------------------------------------
' Last update: 23.10.2018
' ”даление всех классов
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim newFileName As String, tmpFileName As String    ' имена файлов
    Dim curModule As VBComponent                        '
    Dim tmpModule As String
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    newFileName = ThisWorkbook.FullName
    Call unprotectVBE(ThisWorkbook, VBA_PASS)
    For Each curModule In ThisWorkbook.VBProject.VBComponents
        If curModule.Type = vbext_ct_ClassModule Then
            ThisWorkbook.VBProject.VBComponents.remove curModule
        End If
    Next curModule
    ThisWorkbook.Save
    
    tmpModule = Dir(modulesDir())
    While tmpModule <> ""
        If Right(tmpModule, 3) = "cls" Then
            ThisWorkbook.VBProject.VBComponents.Import modulesDir & tmpModule
        End If
        tmpModule = Dir
    Wend
    ThisWorkbook.Save
    ' выполнение обновлений из модул€ "updates"
    Application.ScreenUpdating = SUStatus
    Exit Sub
    
    ' ќбработка ошибок
    ' удаление временных файлов и возврат текущего
errHandler:
    MsgBox "ќшибка обновлени€:" & vbCr & Err.Description
    Err.Clear
    Application.ScreenUpdating = SUStatus
    On Error GoTo 0
        
End Sub

Sub test()
'    Call DeleteCls(True)
End Sub

' ----------------------------------------------------------------------------
' Name: modulesDir function
' Return: string
' Last update: 13.09.2016
' About: путь к каталогу сохранени€ модулей
' ----------------------------------------------------------------------------
Private Function modulesDir() As String
    modulesDir = ThisWorkbook.Path & Application.PathSeparator & "modules" & _
                                                    Application.PathSeparator
End Function
