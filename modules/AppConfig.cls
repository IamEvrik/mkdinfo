VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ��������� ����������
' ----------------------------------------------------------------------------

Private Const cAppVersion As String = "2.46.0"
Private Const cAppName As String = "totalmkd"
Private Const cImportAVRVersion As String = "2.1.0"
Private Const cImportFlatsVersion As String = "1.0.0"
Private Const cImportExpensesVersion As String = "1.0.0"
Private Const cImportSubAccountsVersion As String = "1.0.0"
Private Const cImportPlanSubAccountsVersion As String = "1.0.0"
Private Const cImportSubAccountsMonthVersion As String = "1.0"

Private Const m_templateDir = "templates"
Private Const m_annexDir = "offers"
Private Const m_Razdel3FileName = "������3.xltx"

Private Const m_xmlFullFlatsVersion As Integer = 1
Private Const m_xmlFullFlatsType As String = "full_flats"

Private Const m_serverTmpPath = "\\common\postgres$\"
Private Const m_inServerPath = "e:\exchange\postgres\"

Private Const defaultDBName As String = "totalmkd"
Private Const cVisibleName As String = "�������� ����������"
Private Const cIniFileName = "passport.ini"

Private cDBName As String
Private iniFile As String
Private cGWT As Long
Private cReportFileName As String
Private cBldnPassportFileName As String
Private cDBServer As String


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' Last update: 29.05.2019
' ----------------------------------------------------------------------------
    Dim lLen As Long
    Dim sReturn As String * 255
    
    iniFile = ThisWorkbook.Path & Application.PathSeparator & cIniFileName
    
    lLen = GetPrivateProfileString("GENERAL", "DB", defaultDBName, sReturn, _
                                                                255, iniFile)
    cDBName = Left(sReturn, lLen)

    lLen = GetPrivateProfileString("GENERAL", "GWT", "2", sReturn, _
                                                                255, iniFile)
    cGWT = CLng(Left(sReturn, lLen))
    lLen = GetPrivateProfileString("GENERAL", "REPORT_FILE_NAME", _
                                            "aktYYYY", sReturn, 255, iniFile)
    cReportFileName = Left(sReturn, lLen)
    lLen = GetPrivateProfileString("GENERAL", "PASSPORT_FILE_NAME", _
                                            "�������", sReturn, 255, iniFile)
    cBldnPassportFileName = Left(sReturn, lLen)
    
    lLen = GetPrivateProfileString("CONNECTION", "SERVER", _
                                            "localhost", sReturn, 255, iniFile)
    cDBServer = Left(sReturn, lLen)
End Sub


Property Get dbName() As String
' ----------------------------------------------------------------------------
' �������� ���� ������
' Last update: 11.04.2018
' ----------------------------------------------------------------------------
    dbName = cDBName
End Property


Property Get GWT() As Long
' ----------------------------------------------------------------------------
' ��� ��������
' Last update: 11.04.2018
' ----------------------------------------------------------------------------
    GWT = cGWT
End Property


Property Get AppVersion() As String
' ----------------------------------------------------------------------------
' ������ ����������
' Last update: 11.04.2018
' ----------------------------------------------------------------------------
    AppVersion = cAppVersion
End Property


Property Get appName() As String
' ----------------------------------------------------------------------------
' �������� ����������
' Last update: 11.04.2018
' ----------------------------------------------------------------------------
    appName = cAppName
End Property


Property Get VisibleName() As String
' ----------------------------------------------------------------------------
' ��������� �������� ����������
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    VisibleName = cVisibleName
End Property


Property Get AvrImportVersion() As String
' ----------------------------------------------------------------------------
' ������ ����� ������� ���
' Last update: 07.05.2018
' ----------------------------------------------------------------------------
    AvrImportVersion = cImportAVRVersion
End Property


Property Get FlatsImportVersion() As String
' ----------------------------------------------------------------------------
' ������ ����� ������� ���������� � ����������
' Last update: 25.05.2018
' ----------------------------------------------------------------------------
    FlatsImportVersion = cImportFlatsVersion
End Property


Property Get ExpensesImportVersion() As String
' ----------------------------------------------------------------------------
' ������ ����� ������� ��������� � ��������
' Last update: 29.06.2018
' ----------------------------------------------------------------------------
    ExpensesImportVersion = cImportExpensesVersion
End Property


Property Get IniFileName() As String
' ----------------------------------------------------------------------------
' ���� �� ini-�����
' Last update: 29.06.2018
' ----------------------------------------------------------------------------
    IniFileName = ThisWorkbook.Path & Application.PathSeparator & cIniFileName
End Property


Property Get ReportFileName() As String
' ----------------------------------------------------------------------------
' �������� ����� ��������� ����������
' Last update: 18.10.2018
' ----------------------------------------------------------------------------
    ReportFileName = cReportFileName
End Property


Property Get BldnPassportFileName() As String
' ----------------------------------------------------------------------------
' �������� ����� �������� ����
' Last update: 11.03.2019
' ----------------------------------------------------------------------------
    BldnPassportFileName = cBldnPassportFileName
End Property


Property Get SubAccountsImportVersion() As String
' ----------------------------------------------------------------------------
' ������ ����� ������� ���������� � ���������
' Last update: 08.04.2019
' ----------------------------------------------------------------------------
    SubAccountsImportVersion = cImportSubAccountsVersion
End Property


Property Get PlanSubAccountsImportVersion() As String
' ----------------------------------------------------------------------------
' ������ ����� ������� ���������� � �������� ������������ �� ��������
' Last update: 17.06.2019
' ----------------------------------------------------------------------------
    PlanSubAccountsImportVersion = cImportPlanSubAccountsVersion
End Property


Property Get SubAccountsMonthVersion() As String
' ----------------------------------------------------------------------------
' ������ ����� ������� �������� ���������� � ���������
' Last update: 11.07.2019
' ----------------------------------------------------------------------------
    SubAccountsMonthVersion = cImportSubAccountsMonthVersion
End Property


Property Get ServerTmpPath() As String
' ----------------------------------------------------------------------------
' ���� � ��������� �� �������
' 09.08.2021
' ----------------------------------------------------------------------------
    If cDBServer = "localhost" Then
        ServerTmpPath = "d:\temp\"
    Else
        ServerTmpPath = m_serverTmpPath
    End If
End Property


Property Get InServerTmpPath() As String
' ----------------------------------------------------------------------------
' ���� � ��������� ��� ������ �������
' 09.08.2021
' ----------------------------------------------------------------------------
    If cDBServer = "localhost" Then
        InServerTmpPath = "d:\temp\"
    Else
        InServerTmpPath = m_inServerPath
    End If
End Property


Property Get DBServer() As String
' ----------------------------------------------------------------------------
' ����� �������
' Last update: 29.05.2019
' ----------------------------------------------------------------------------
    DBServer = cDBServer
End Property


Property Get xmlVersion(xmlHeaderType As XmlHeaderTypeEnum) As Integer
' ----------------------------------------------------------------------------
' ������ xml-����� ��� ��������
' 12.08.2021
' ----------------------------------------------------------------------------
    Select Case (xmlHeaderType)
    
        Case XmlHeaderTypeEnum.xhtFlats
            xmlVersion = m_xmlFullFlatsVersion
            
    End Select
End Property


Property Get xmlType(xmlHeaderType As XmlHeaderTypeEnum) As String
' ----------------------------------------------------------------------------
' ��� xml-����� ��� ��������
' 12.08.2021
' ----------------------------------------------------------------------------
    Select Case (xmlHeaderType)
    
        Case XmlHeaderTypeEnum.xhtFlats
            xmlType = m_xmlFullFlatsType
            
    End Select
End Property


Property Get Razdel3FileName() As String
' ----------------------------------------------------------------------------
' ���� � ������� ������3
' 20.10.2021
' ----------------------------------------------------------------------------
    Razdel3FileName = ThisWorkbook.Path & Application.PathSeparator & _
                m_templateDir & Application.PathSeparator & _
                m_annexDir & Application.PathSeparator & m_Razdel3FileName
End Property


Public Function ValidVersion(curVersion As String) As Boolean
' ----------------------------------------------------------------------------
' �������� ������ ����
' Last update: 14.06.2019
' ----------------------------------------------------------------------------
    Dim dotCount As Integer, i As Integer
    Dim cApp As String, cBase As String
    Const USED_DOT As Integer = 2
    
    ' ��������� �� ������ ���������� � ���� ������ ������ ����
    ' (������ � ������������� ������� ������ ���������� ����� �����)
    dotCount = 0
    For i = 1 To Len(cAppVersion)
        If mid(cAppVersion, i, 1) = "." Then
            dotCount = dotCount + 1
            If dotCount = USED_DOT Then Exit For
        End If
        cApp = cApp & mid(cAppVersion, i, 1)
    Next i
    
    dotCount = 0
    For i = 1 To Len(curVersion)
        If mid(curVersion, i, 1) = "." Then
            dotCount = dotCount + 1
            If dotCount = USED_DOT Then Exit For
        End If
        cBase = cBase & mid(curVersion, i, 1)
    Next i
    
    If StrComp(cApp, cBase, vbTextCompare) = 0 Then
        ValidVersion = True
    Else
        ValidVersion = False
    End If
            
End Function
