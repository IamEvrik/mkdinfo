VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCounterModel 
   Caption         =   "Модель прибора учёта"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "FormCounterModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCounterModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As counter_model


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' Активация формы - заполнение полей если изменение
' Last update: 13.05.2019
' ----------------------------------------------------------------------------
    If curItem Is Nothing Then
        Me.Caption = "Добавление модели приборов учёта"
        Me.BtnSave.Caption = "Добавить"
    Else
        Me.Caption = "Изменение модели приборов учёта"
        Me.TextBoxName = curItem.Name
        Me.TextBoxCI = curItem.CalibrationInterval
        Me.CBHasDTI = curItem.HasDTI
        Me.BtnSave.Caption = "Сохранить"
    End If
    Me.Caption = Me.Caption & ". " & AppConfig.DBServer
End Sub


Private Sub BtnSave_Click()
' ----------------------------------------------------------------------------
' Обработка кнопки сохранения
' Last update: 13.05.2019
' ----------------------------------------------------------------------------
    If formFill Then
        If curItem Is Nothing Then
            Set curItem = New counter_model
            curItem.create Name:=Trim(Me.TextBoxName), _
                            HasDTI:=Me.CBHasDTI, _
                            CalibrationInterval:=Trim(Me.TextBoxCI.Value)
        Else
            curItem.update newName:=Trim(Me.TextBoxName), _
                            newHasDTI:=Me.CBHasDTI, _
                            newCI:=Trim(Me.TextBoxCI.Value)
        End If
        Set curItem = Nothing
        Unload Me
    End If
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' Обработка кнопки отмены
' Last update: 13.05.2019
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Function formFill() As Boolean
' ----------------------------------------------------------------------------
' Проверка заполнения полей
' Last update: 13.05.2019
' ----------------------------------------------------------------------------
    formFill = False
    If Len(Trim(Me.TextBoxCI)) + Len(Trim(Me.TextBoxName)) > 0 Then
        If longValue(Me.TextBoxCI) > 0 Then
            formFill = True
        Else
            MsgBox "Неверное значение 'Межповерочный интервал'", _
                                        vbExclamation, "Ошибка при проверке"
        End If
    Else
        MsgBox "Заполнены не все поля", vbExclamation, "Ошибка при проверке"
    End If
End Function
