Attribute VB_Name = "ErrorNumbers"
' �������������� ������
Public Const ERROR_NOT_VALID_VALUE = 60001          ' �������� �� �������������
Public Const ERROR_OBJECT_NOT_SET = 60002           ' ������ �� ����������
Public Const ERROR_INDEX_OUT_OF_DICT = 60003        ' ������ ������� �� ������� �����������
Public Const ERROR_FILE_NOT_FOUND = 60004           ' ���� �� ������
Public Const ERROR_SQLITE_NON_QUERY = 60005         ' ������� sqlite3nonquery ��������� ��������
Public Const ERROR_NOT_VALID_VERSION = 60006        ' ������������ ������ ����� � ���� ������
Public Const ERROR_OBJECT_NOT_VALID = 60007         ' � ������� ������� ������������ ������

Public Const DB_ERROR_OBJECT_HAS_CHILDREN = "23503"

Public Const ERROR_NOT_UNIQUE = 60008               ' ������������ ��������
Public Const ERROR_OBJECT_HAS_CHILDREN = 60009      ' ������ ������� ������, �.�. ���� ��������
Public Const ERROR_NOT_DELETE = 99003               ' ������ ����� ������� ������

Public Const ERROR_PRIVILEGE_NOT_GRANTED = 60010    ' �� ������� ����
Public Const ERROR_INVALID_LOAD_FILE_VERSION = 60011 ' ������������ ������ ������������ �����


Public Function errorHasChildren(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' �������� ������ ������ �� ������� �������� ����������� fk
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    errorHasChildren = (InStr(1, errDesc, DB_ERROR_OBJECT_HAS_CHILDREN, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorNotUnique(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' �������� ������ ������ �� ������� �������� ������������
' 19.10.2022
' ----------------------------------------------------------------------------
    errorNotUnique = (InStr(1, errDesc, "unique", _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorStopDelete(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' �������� ������ ������ �� ������� �������� ������������� ��������
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    errorStopDelete = (InStr(1, errDesc, ERROR_NOT_DELETE, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorHasNoPrivilegies(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' �������� ������ ������ �� ������� �������� ���������� ����
' Last update: 25.09.2019
' ----------------------------------------------------------------------------
    errorHasNoPrivilegies = (InStr(1, errDesc, ERROR_PRIVILEGE_NOT_GRANTED, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorHasNoValues(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' �������� ������ ������ �� ������� �������� "��� ������"
' Last update: 01.06.2021
' ----------------------------------------------------------------------------
    errorHasNoValues = (InStr(1, errDesc, ERROR_OBJECT_NOT_SET, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function getErrorText(currentError As ErrObject) As String
' ----------------------------------------------------------------------------
' ������ ������ ������ ������ ��� ������������� ������
' 25.10.2022
' ----------------------------------------------------------------------------
    If errorHasChildren(Err.Description) Then
        getErrorText = "���� ����������� �������, �������� ���������"
    ElseIf errorHasNoValues(Err.Number) Then
        getErrorText = "������ �� �����"
    ElseIf errorHasNoPrivilegies(Err.Description) Then
        getErrorText = "� ��� �� ���������� ����"
    ElseIf errorNotUnique(Err.Description) Then
        getErrorText = "� ���� ��� ���� ������ � ������ �����������"
    Else
        getErrorText = Err.Number & " " & Err.Description
    End If
End Function
