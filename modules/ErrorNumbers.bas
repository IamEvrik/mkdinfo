Attribute VB_Name = "ErrorNumbers"
' идентификаторы ошибок
Public Const ERROR_NOT_VALID_VALUE = 60001          ' Значение не соответствует
Public Const ERROR_OBJECT_NOT_SET = 60002           ' Объект не существует
Public Const ERROR_INDEX_OUT_OF_DICT = 60003        ' индекс выходит за границы справочника
Public Const ERROR_FILE_NOT_FOUND = 60004           ' файл не найден
Public Const ERROR_SQLITE_NON_QUERY = 60005         ' функция sqlite3nonquery выполнена неудачно
Public Const ERROR_NOT_VALID_VERSION = 60006        ' неправильная версия файла и базы данных
Public Const ERROR_OBJECT_NOT_VALID = 60007         ' в функцию передан неправильный объект

Public Const DB_ERROR_OBJECT_HAS_CHILDREN = "23503"

Public Const ERROR_NOT_UNIQUE = 60008               ' дублирование названия
Public Const ERROR_OBJECT_HAS_CHILDREN = 60009      ' нельзя удалить объект, т.к. есть дочерние
Public Const ERROR_NOT_DELETE = 99003               ' данный пункт удалять нельзя

Public Const ERROR_PRIVILEGE_NOT_GRANTED = 60010    ' не хватает прав
Public Const ERROR_INVALID_LOAD_FILE_VERSION = 60011 ' неправильная версия загружаемого файла


Public Function errorHasChildren(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' проверка текста ошибки на наличие признака ограничения fk
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    errorHasChildren = (InStr(1, errDesc, DB_ERROR_OBJECT_HAS_CHILDREN, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorNotUnique(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' проверка текста ошибки на наличие признака уникальности
' 19.10.2022
' ----------------------------------------------------------------------------
    errorNotUnique = (InStr(1, errDesc, "unique", _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorStopDelete(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' проверка текста ошибки на наличие признака невозможности удаления
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    errorStopDelete = (InStr(1, errDesc, ERROR_NOT_DELETE, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorHasNoPrivilegies(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' проверка текста ошибки на наличие признака недостатка прав
' Last update: 25.09.2019
' ----------------------------------------------------------------------------
    errorHasNoPrivilegies = (InStr(1, errDesc, ERROR_PRIVILEGE_NOT_GRANTED, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function errorHasNoValues(errDesc As String) As Boolean
' ----------------------------------------------------------------------------
' проверка текста ошибки на наличие признака "нет данных"
' Last update: 01.06.2021
' ----------------------------------------------------------------------------
    errorHasNoValues = (InStr(1, errDesc, ERROR_OBJECT_NOT_SET, _
                                                        vbBinaryCompare) > 0)
End Function


Public Function getErrorText(currentError As ErrObject) As String
' ----------------------------------------------------------------------------
' выдача своего текста ошибки для отслеживаемых ошибок
' 25.10.2022
' ----------------------------------------------------------------------------
    If errorHasChildren(Err.Description) Then
        getErrorText = "Есть подчиненные объекты, действие запрещено"
    ElseIf errorHasNoValues(Err.Number) Then
        getErrorText = "Объект не задан"
    ElseIf errorHasNoPrivilegies(Err.Description) Then
        getErrorText = "У Вас не достаточно прав"
    ElseIf errorNotUnique(Err.Description) Then
        getErrorText = "В базе уже есть запись с такими параметрами"
    Else
        getErrorText = Err.Number & " " & Err.Description
    End If
End Function
