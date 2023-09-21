Attribute VB_Name = "WinRegFunctions"
Option Explicit

' ----------------------------------------------------------------------------
' �������� �� ���������
' ----------------------------------------------------------------------------
Private Const PREFETCH_GWT = "2"
Private Const INI_FILE_NAME = "passport.ini"


' #TODO: ������ ����������� ����-������, ����� ���� � ������, ������ �� ������
Public Sub saveUseWorkDate(newValue As Long)
' ----------------------------------------------------------------------------
' ���������� ���������� ������� ����� � ������
' Last update: 23.05.2016
' ----------------------------------------------------------------------------
    SaveSetting appName, "works", "lastdate", CStr(newValue)
End Sub


' #TODO: ����� ����������� ����-������ �� ����, ����� � ������ ������
Public Function getUseWorkDate() As Long
' ----------------------------------------------------------------------------
' ��������� ���������� ������� ����� �� �������, ���� ��� �� ������,
'           �� ������������ NOTVALUE
' Last update: 23.05.2016
' ----------------------------------------------------------------------------
    getUseWorkDate = CLng(getSetting(appName, "works", "lastdate", _
                                                            CStr(NOTVALUE)))
End Function


Public Function getPrefetchWork() As String
' ----------------------------------------------------------------------------
' ��������� ������������� ������
' Last update: 28.09.2016
' ----------------------------------------------------------------------------
    getPrefetchWork = ReadIniFile("GWT", PREFETCH_GWT, "GENERAL", _
                ThisWorkbook.Path & Application.PathSeparator & INI_FILE_NAME)
End Function
