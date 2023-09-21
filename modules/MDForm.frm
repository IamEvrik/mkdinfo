VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MDForm 
   Caption         =   "Список муниципальных образований"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   OleObjectBlob   =   "MDForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MDForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As address_md_class


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' Активация формы - заполнение полей если изменение
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Me.TextBoxHeadPosition.Enabled = False
    If curItem Is Nothing Then
        Me.Caption = "Добавление муниципального образования"
        Me.BtnSave.Caption = "Добавить"
    Else
        Me.Caption = "Изменение муниципального образования"
        Me.TextBoxName = curItem.Name
        Me.TextBoxHead = curItem.Head
        Me.BtnSave.Caption = "Сохранить"
        Me.CheckBoxHeadPosition = curItem.HasHeadPosition
    End If
    Me.Caption = Me.Caption & ". " & AppConfig.DBServer
End Sub


Private Sub CheckBoxHeadPosition_Change()
'-----------------------------------------------------------------------------
' изменение доступности поля ввод должности главы
' Last update: 15.02.2018
'-----------------------------------------------------------------------------
    If Me.CheckBoxHeadPosition Then
        Me.TextBoxHeadPosition.Value = curItem.HeadPosition
    Else
        Me.TextBoxHeadPosition.Value = ""
    End If
    Me.TextBoxHeadPosition.Enabled = Me.CheckBoxHeadPosition.Value
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' уничтожение формы
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Set curItem = Nothing
End Sub


Private Sub BtnSave_Click()
' ----------------------------------------------------------------------------
' обработка кнопки добавления
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Call process
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' закрытие формы по кнопке отмена
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub process()
' ----------------------------------------------------------------------------
' добавление/изменения
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Dim addFlag As Boolean
    
    On Error GoTo errHandler

    If formNotFill Then
        MsgBox "Заполнены не все поля!", vbInformation + vbOKOnly, "Ошибка"
        GoTo cleanHandler
    End If
    
    addFlag = False
    If curItem Is Nothing Then
        Set curItem = New address_md_class
        addFlag = True
    End If
    
    curItem.update newName:=Me.TextBoxName, _
                    newHead:=Me.TextBoxHead, _
                    newHeadPosition:=Me.TextBoxHeadPosition, _
                    addNew:=addFlag
    Unload Me
    GoTo cleanHandler
    
errHandler:
    Dim errMsg As String
    If errorHasNoPrivilegies(Err.Description) Then
        errMsg = "Не хватает прав"
    ElseIf errorNotUnique(Err.Description) Then
        errMsg = "Данное значение уже существует"
    Else
        errMsg = Err.Description
    End If
    MsgBox errMsg, vbCritical, "Ошибка"
    
cleanHandler:
End Sub







Private Sub process1(addFlag As Boolean)
'-----------------------------------------------------------------------------
' изменение/добавление в базу
' Last update: 22.03.2018
'-----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Or addFlag Then
        If formNotFill Then
            MsgBox "Заполнены не все поля!", vbInformation, "Ошибка"
        Else
            On Error GoTo errHandler
            curItem.update newName:=Trim(Me.TextBoxName.Value), _
                        newHead:=Trim(Me.TextBoxHead.Value), _
                        newHeadPosition:=Trim(Me.TextBoxHeadPosition.Value), _
                        addNew:=addFlag
         
            Call clearTextBox
            Call reloadListView
        End If
    End If
        
    GoTo cleanHandler

errHandler:
    If Err.Number = ERROR_NOT_UNIQUE Then
        MsgBox Err.Description, vbInformation, "Ошибка"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

cleanHandler:
End Sub


Private Function formNotFill() As Boolean
'-----------------------------------------------------------------------------
' проверка заполнения обязательных полей
' Last update: 22.03.2017
'-----------------------------------------------------------------------------
    formNotFill = (Len(Trim(Me.TextBoxHead.Value)) = 0 Or _
                    Len(Trim(Me.TextBoxName.Value)) = 0)
End Function


