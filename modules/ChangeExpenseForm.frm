VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeExpenseForm 
   Caption         =   "Изменение расходов"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   OleObjectBlob   =   "ChangeExpenseForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeExpenseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public expenseId As Long


Private Sub BtnSave_Click()
' ----------------------------------------------------------------------------
' сохранение изменений
' Last update: 10.04.2019
' ----------------------------------------------------------------------------
    If dblValue(Me.TextBoxPrice.Value) = NOTVALUE Or _
                        dblValue(Me.TextBoxPlanSum.Value) = NOTVALUE Or _
                        dblValue(Me.TextBoxFactSum.Value) = NOTVALUE Then
        MsgBox "Проверьте правильность ввода", vbCritical, "Ошибка"
        Exit Sub
    Else
        Dim tmp As New expense
        tmp.change Id:=expenseId, _
                    planSum:=dblValue(Me.TextBoxPlanSum.Value), _
                    factSum:=dblValue(Me.TextBoxFactSum.Value), _
                    price:=dblValue(Me.TextBoxPrice.Value)
        Set tmp = Nothing
        Call CloseMe(True)
    End If
End Sub


Private Sub BtnExit_Click()
' ----------------------------------------------------------------------------
' выход
' Last update: 06.07.2018
' ----------------------------------------------------------------------------
    Call CloseMe
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' при активации фокус на фактической сумме
' Last update: 10.04.2019
' ----------------------------------------------------------------------------
    Me.TextBoxFactSum.SelStart = 0
    Me.TextBoxFactSum.SelLength = Len(Me.TextBoxFactSum.text)
    Me.TextBoxFactSum.SetFocus
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' запрет закрытия формы крестиком,
'                    т.к. после этого некорректно работает показ формы МКД
' Last update: 05.07.2018
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


Private Sub CloseMe(Optional updateBldn As Boolean = False)
' ----------------------------------------------------------------------------
' закрытие формы и обновление структуры на форме дома
' Last update: 06.07.2018
' ----------------------------------------------------------------------------
    Unload Me
    If updateBldn Then BuildingForm.reloadExpenseList
    BuildingForm.ListViewExpenses.SetFocus
End Sub
