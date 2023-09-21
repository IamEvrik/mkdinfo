Attribute VB_Name = "WinRegFunctions"
Option Explicit

' ----------------------------------------------------------------------------
' Значения по умолчанию
' ----------------------------------------------------------------------------
Private Const PREFETCH_GWT = "2"
Private Const INI_FILE_NAME = "passport.ini"


' #TODO: должно сохраняться куда-нибудь, пусть даже в инишку, реестр не гадить
Public Sub saveUseWorkDate(newValue As Long)
' ----------------------------------------------------------------------------
' сохранение последнего периода работ в реестр
' Last update: 23.05.2016
' ----------------------------------------------------------------------------
    SaveSetting appName, "works", "lastdate", CStr(newValue)
End Sub


' #TODO: пусть сохраняется куда-нибудь на лист, нефиг в реестр гадить
Public Function getUseWorkDate() As Long
' ----------------------------------------------------------------------------
' получение последнего периода работ из реестра, если еще не задано,
'           то возвращается NOTVALUE
' Last update: 23.05.2016
' ----------------------------------------------------------------------------
    getUseWorkDate = CLng(getSetting(appName, "works", "lastdate", _
                                                            CStr(NOTVALUE)))
End Function


Public Function getPrefetchWork() As String
' ----------------------------------------------------------------------------
' получение препочитаемой работы
' Last update: 28.09.2016
' ----------------------------------------------------------------------------
    getPrefetchWork = ReadIniFile("GWT", PREFETCH_GWT, "GENERAL", _
                ThisWorkbook.Path & Application.PathSeparator & INI_FILE_NAME)
End Function
