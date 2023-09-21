Attribute VB_Name = "APImodule"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
#Else
    Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
#End If
Const SM_CXSCREEN = 0, SM_CYSCREEN = 1

'API Declaration in General Declarations
#If VBA7 Then
    Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#Else
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

'API Constants
Const SET_COLUMN_WIDTH As Long = 4126
Const AUTOSIZE_USEHEADER As Long = -2

' create directory
#If VBA7 Then
    Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                                     (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
#Else
    Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                                         (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
#End If

' ���������� �� ini-�����
#If VBA7 Then
    Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If
                                          
' ������ � ini-����
#If VBA7 Then
    Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                           (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                           ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                           (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                           ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Sub one()
    Dim x, y
    x = GetSystemMetrics(SM_CXSCREEN)
    y = GetSystemMetrics(SM_CYSCREEN)
    MsgBox "���������� ������ " & x & " * " & y
'    MOForm.Width = X * 0.75
'    MOForm.Height = Y * 0.75
'    MOForm.Show
End Sub


'Sub To Resize
Sub AppNewAutosizeColumns(ByVal TargetListView As ListView)

    Const SET_COLUMN_WIDTH As Long = 4126
    Const AUTOSIZE_USEHEADER As Long = -1

    Dim lngColumn As Long

    For lngColumn = 0 To (TargetListView.ColumnHeaders.count - 1)

        Call SendMessage(TargetListView.hwnd, _
            SET_COLUMN_WIDTH, _
            lngColumn, _
            AUTOSIZE_USEHEADER)

    Next lngColumn

End Sub

'sub to create folder
Sub CreateFolders(ByVal folderPath As String)
    SHCreateDirectoryEx Application.hwnd, folderPath, 0&
End Sub

' ������ ������ � ����
Public Sub WriteIniFile(ByVal sName$, ByVal val$, ByVal sPart$, ByVal FilePath$)
    ' ������� ���� � ini ����� FilePath$ ������ sPart$ (���� ������� ��� - �� ��������),
    ' � ��������� � ���� �������� � ������ sName$ � ��������� val
    Dim intRet As Integer: intRet = WritePrivateProfileString(sPart, sName, val, FilePath)
    'If intRet <> 1 Then '��������� ����������'(�������� ���������� ������)
End Sub
 
' ��������� ������ �� ini �����
Public Function ReadIniFile(ByVal sName$, ByVal DefVal$, ByVal sPart$, ByVal FilePath$) As String
    ' ������� ���� � ini ����� FilePath$ ������ sPart$,
    ' � ������ �� ���� �������� �� ��������� � ������ sName$
    ' ���� ����� �������� �� ������, ������������ �������� �� ��������� DefVal$

    Const strNoValue As String = ""
    Dim intRet As Integer    ' ����� ������������ ������ (�������� GetPrivateProfileString)
    Dim strRet As String    ' ������������ ������
    ' �������� �������� �� ����� - ���� ��� ���, ����� ��������� 3� �������� = strNoValue
    strRet = String(255, Chr(0)): intRet = GetPrivateProfileString(sPart, sName, strNoValue, strRet, 255, FilePath)
    strRet = Left$(strRet, intRet)
    ' ����������, ���� ������� �������� ��� ��� (���� ���������� strNoValue �� = ���)
    If strRet = strNoValue Then strRet = DefVal   ' �������� �� ���� ������� - ���������� �������� �� ���������
    ReadIniFile = strRet
End Function
