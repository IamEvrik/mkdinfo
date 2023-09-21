VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
   OleObjectBlob   =   "BaseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public curItem As base_form_class
Public isChange As Boolean

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonSave_Click()
    On Error GoTo errHandler
    
    If Not curItem.isFormFill(Me) Then
        MsgBox "Не заполнены все обязательные поля", vbExclamation, "Внимание"
    Else
        curItem.update Me, isChange
    End If
    
errHandler:
    If Err.Number <> 0 Then Me.LabelError.Caption = Err.Description
End Sub
