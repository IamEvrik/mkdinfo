VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpenseNamesForm 
   Caption         =   "Названия статей расходов по домам"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "ExpenseNamesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExpenseNamesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & " " & AppConfig.DBServer
    Call reloadComboBox(rcmExpenseItems, Me.ComboBoxExpenseItems)
End Sub


Private Sub ComboBoxExpenseItems_Change()
' ----------------------------------------------------------------------------
' Заполнение данных при изменении статьи расходов
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxExpenseItems.ListIndex > -1 Then
        Me.LabelName1.Caption = expense_items( _
                                CStr(Me.ComboBoxExpenseItems.Value)).Name1
        Me.LabelName2.Caption = expense_items( _
                                CStr(Me.ComboBoxExpenseItems.Value)).Name2
        Call reloadComboBox(rcmBldnExpenseName, Me.ListBoxName1, _
                    initValue:=Me.ComboBoxExpenseItems.Value, initValue2:=1)
        Call reloadComboBox(rcmBldnExpenseName, Me.ListBoxName2, _
                    initValue:=Me.ComboBoxExpenseItems.Value, initValue2:=2)
    End If
End Sub


Private Sub Btn1to2_Click()
' ----------------------------------------------------------------------------
' перенос дома из списка 1 в список 2
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To Me.ListBoxName1.ListCount - 1
        If Me.ListBoxName1.Selected(i) Then
            Call updateBldnExpenseName(Me.ListBoxName1.Value, _
                                            Me.ComboBoxExpenseItems.Value, 2)
        End If
    Next i
    Call MoveListBoxElements(Me.ListBoxName1, Me.ListBoxName2)
End Sub


Private Sub Btn2to1_Click()
' ----------------------------------------------------------------------------
' перенос дома из списка 2 в список 1
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To Me.ListBoxName1.ListCount - 1
        If Me.ListBoxName2.Selected(i) Then
            Call updateBldnExpenseName(Me.ListBoxName2.Value, _
                                            Me.ComboBoxExpenseItems.Value, 1)
        End If
    Next i
    Call MoveListBoxElements(Me.ListBoxName2, Me.ListBoxName1)
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' закрытие формы
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Unload Me
End Sub
