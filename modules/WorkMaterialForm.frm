VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkMaterialForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "WorkMaterialForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkMaterialForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public curWork As work_maintenance
Public materialIdx As Long


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' инициализация формы, заполнение полей
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    Call reloadComboBox(rcmWorkMaterialTypes, Me.ComboBoxMaterials, _
                                                                defValue:=0)
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы, заполнение полей
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    If materialIdx <> 0 Then
        With curWork.Materials(materialIdx)
            Call selectComboBoxValue(Me.ComboBoxMaterials, .MaterialId)
            Me.TextBoxCost = .MaterialCost
            Me.TextBoxCount = .MaterialCount
            Me.TextBoxNote = .MaterialNote
            Me.TextBoxSI = .MaterialSi
        End With
    End If
    Me.Caption = Me.Caption & ". Сервер " & AppConfig.DBServer
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' запрет закрытия формы крестиком
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' сохранение
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    If formFill Then
        Dim lIdx As Long
        If materialIdx <> 0 Then
            lIdx = materialIdx
        Else
            curWork.Materials.add New works_materials
            lIdx = curWork.Materials.count
        End If
        With curWork.Materials(lIdx)
            .MaterialCost = dblValue(Me.TextBoxCost)
            .MaterialCount = dblValue(Me.TextBoxCount)
            .MaterialId = Me.ComboBoxMaterials.Value
            .MaterialNote = Me.TextBoxNote
            .MaterialSi = Me.TextBoxSI
        End With
        Unload Me
    Else
        MsgBox "Заполнены не все поля", , "Внимание"
    End If
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' выход без сохранения
' Last update: 18.10.2019
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Function formFill() As Boolean
' ----------------------------------------------------------------------------
' проверка заполнения полей
' Last update: 16.10.2019
' ----------------------------------------------------------------------------
    formFill = (dblValue(Me.TextBoxCost.Value) > 0 And _
                                        dblValue(Me.TextBoxCount.Value) > 0)
End Function
