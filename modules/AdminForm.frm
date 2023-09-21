VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdminForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   OleObjectBlob   =   "AdminForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' Заголовок формы
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Me.Caption = "Администрирование. Сервер " & AppConfig.DBServer
End Sub


Private Sub ButtonUsers_Click()
' ----------------------------------------------------------------------------
' Показ формы списка пользователей
' Last update: 26.09.2019
' ----------------------------------------------------------------------------
    Call RunUserListForm
End Sub


Private Sub ButtonUserRoles_Click()
' ----------------------------------------------------------------------------
' Показ формы прав доступа пользователей
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Call RunUserRolesForm
End Sub


Private Sub ButtonRolesAccess_Click()
' ----------------------------------------------------------------------------
' Показ формы прав групп
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Call RunUserRolesAccessForm
End Sub


Private Sub BtnCreateBackup_Click()
' ----------------------------------------------------------------------------
' Создание архива
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Call createBackupPG
End Sub


Private Sub BtnExportModules_Click()
' ----------------------------------------------------------------------------
' Выгрузка модулей
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Call ExportModules
End Sub

