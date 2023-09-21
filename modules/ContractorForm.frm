VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContractorForm 
   Caption         =   "Список подрядчиков"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   OleObjectBlob   =   "ContractorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContractorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As contractor_class


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' Активация формы - заполнение полей если изменение
' 17.09.2021
' ----------------------------------------------------------------------------
    If curItem Is Nothing Then
        Me.Caption = "Добавление подрядной организации"
        Me.BtnSave.Caption = "Добавить"
        Me.CheckBoxUsing = True
        Me.CheckBoxUsing.Enabled = False
    Else
        Me.Caption = "Изменение подрядной организации"
        Me.TextBoxName = curItem.Name
        Me.TextBoxDirector = curItem.Director
        Me.TextBoxDirPosition = curItem.DirectorPosition
        Me.CheckBoxBldnContractor = curItem.BldnContractor
        Me.CheckBoxUsing = curItem.isUsing
        Me.BtnSave.Caption = "Сохранить"
    End If
    Me.Caption = Me.Caption & ". " & AppConfig.DBServer
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' уничтожение формы
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Set curItem = Nothing
End Sub


Private Sub BtnSave_Click()
' ----------------------------------------------------------------------------
' обработка кнопки добавления подрядчика
' Last update: 29.05.2019
' ----------------------------------------------------------------------------
    Call process
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' закрытие формы по кнопке отмена
' Last update: 29.05.2019
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub process()
' ----------------------------------------------------------------------------
' добавление/изменения подрядчика
' 17.09.2021
' ----------------------------------------------------------------------------
    Dim addFlag As Boolean
    
    On Error GoTo errHandler

    If formNotFill Then
        MsgBox "Заполнены не все поля!", vbInformation + vbOKOnly, "Ошибка"
        GoTo cleanHandler
    End If
    
    addFlag = False
    If curItem Is Nothing Then
        Set curItem = New contractor_class
        addFlag = True
    End If
    
    curItem.update newName:=Me.TextBoxName.Value, _
                    newDirector:=Me.TextBoxDirector.Value, _
                    newBldnStatus:=Me.CheckBoxBldnContractor.Value, _
                    newDirectorPosition:=Me.TextBoxDirPosition.Value, _
                    newIsUsing:=Me.CheckBoxUsing, _
                    addNew:=addFlag
    Unload Me
    GoTo cleanHandler
        
errHandler:
    If Err.Number = ERROR_NOT_UNIQUE Then
        MsgBox Err.Description, vbInformation, "Ошибка"
    Else
        MsgBox Err.Number & Err.Source & Err.Description, vbExclamation, "Ошибка"
    End If
        
cleanHandler:
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' проверка на заполнение необходимых полей
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    formNotFill = (StrComp(Trim(Me.TextBoxName.Value), "") = 0)
End Function
