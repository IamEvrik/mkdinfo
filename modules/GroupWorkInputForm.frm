VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GroupWorkInputForm 
   Caption         =   "Ввод работ на группу домов"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12900
   OleObjectBlob   =   "GroupWorkInputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GroupWorkInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
'-----------------------------------------------------------------------------
' инициализация формы, заполнение списков
' Last update:17.04.2018
'-----------------------------------------------------------------------------
    Call reloadComboBox(rcmContractor, Me.ComboBoxContractor)
    Call reloadComboBox(rcmGWT, Me.ComboBoxGlobalWorkType)
    Call reloadComboBox(rcmWorkType, Me.ComboBoxWorkType)
    Call reloadComboBox(rcmTerm, Me.ComboBoxTerms)
    Call reloadComboBox(rcmListBldnAddressId, Me.ListBoxAvailable)
    Call reloadComboBox(rcmMC, Me.ComboBoxMC)
    Call reloadComboBox(rcmFSources, Me.ComboBoxFSource)
    Call reloadComboBox(rcmListBldnAddressId, Me.ListBoxAvailable)
    Me.ComboBoxTerms.ListIndex = Me.ComboBoxTerms.ListCount - 1
    Me.ComboBoxWorkKind.Enabled = False
End Sub


Private Sub ButtonClearSelection_Click()
'-----------------------------------------------------------------------------
' Очистка списка домов
' Last update: 29.06.2017
'-----------------------------------------------------------------------------
    Me.ListBoxSelected.Clear
    Call reloadComboBox(rcmListBldnAddressId, Me.ListBoxAvailable)
End Sub


Private Sub ButtonSaveWorks_Click()
'-----------------------------------------------------------------------------
' Сохранение работы
' Last update: 17.04.2018
'-----------------------------------------------------------------------------
    Dim curWorkSum As Double
    Dim i As Long
    Dim curWork As New work_class
    Dim errMsg As String
    
    If Me.ComboBoxContractor.ListIndex > -1 And _
                Me.ComboBoxGlobalWorkType.ListIndex > -1 And _
                Me.ComboBoxTerms.ListIndex > -1 And _
                Me.ComboBoxWorkKind.ListIndex > -1 And _
                Me.ComboBoxWorkType.ListIndex > -1 And _
                Me.TextBoxDogovor.Value <> "" And _
                Me.TextBoxSI.Value <> "" And _
                Me.TextBoxVolume.Value <> "" And _
                Me.TextBoxSum.Value <> "" And _
                Me.ComboBoxMC.ListIndex > -1 And _
                Me.ComboBoxFSource.ListIndex > -1 Then
        If Me.ListBoxSelected.ListCount = 0 Then
            MsgBox "Не выбран ни один дом", vbOKOnly + vbExclamation, _
                                                                "Выберите МКД"
            Exit Sub
        End If
        curWorkSum = dblValue(Me.TextBoxSum.Value)
        If curWorkSum = NOTVALUE Then
            MsgBox "Неверно введена сумма работы", vbCritical + vbOKOnly, _
                                                        "Ошибка сохранения"
            Exit Sub
        End If
        
        
        On Error GoTo errHandler
        
        For i = 0 To Me.ListBoxSelected.ListCount - 1
            Call curWork.create( _
                    BldnId:=Me.ListBoxSelected.list(i, ComboColumns.ccId), _
                    gwtId:=Me.ComboBoxGlobalWorkType.Value, _
                    workKindID:=Me.ComboBoxWorkKind.Value, _
                    WorkDate:=Me.ComboBoxTerms.Value, _
                    workSum:=CCur(curWorkSum), _
                    Si:=Me.TextBoxSI.Value, _
                    workVolume:=Me.TextBoxVolume.Value, _
                    workNote:=Me.TextBoxNote.Value, _
                    contractorId:=Me.ComboBoxContractor.Value, _
                    mcId:=Me.ComboBoxMC.Value, _
                    Dogovor:=Me.TextBoxDogovor.Value, _
                    PrintFlag:=Me.CheckBoxPrintFlag, _
                    financeSource:=Me.ComboBoxFSource.Value)
        Next i
        GoTo cleanHandler
        
errHandler:
        GoTo cleanHandler
    
cleanHandler:
        MsgBox "Работы введены успешно"
    Else
        MsgBox "Заполнены не все поля", vbOKOnly + vbExclamation, _
                                        "Проверьте правильность заполнения"
    End If
End Sub

Private Sub ComboBoxWorkType_Change()
'-----------------------------------------------------------------------------
' заполнение видов работ при изменении типа
' Last update: 29.06.2017
'-----------------------------------------------------------------------------
    If Me.ComboBoxWorkType.ListIndex > -1 Then
        Call reloadComboBox(rcmWorkKind, Me.ComboBoxWorkKind, _
                                initValue:=CLng(Me.ComboBoxWorkType.Value))
        Me.ComboBoxWorkKind.Enabled = True
    End If
End Sub


Private Sub ButtonAddBldn_Click()
'-----------------------------------------------------------------------------
' добавление адреса в выбранные
' Last update: 29.06.2017
'-----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxAvailable, Me.ListBoxSelected)
End Sub


Private Sub ButtonRemoveBldn_Click()
'-----------------------------------------------------------------------------
' удаление адреса из выбранных
' Last update: 29.06.2017
'-----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxSelected, Me.ListBoxAvailable)
End Sub


Private Sub ButtonCancel_Click()
'-----------------------------------------------------------------------------
' нажатие кнопки "Отмена", закрытие формы
' Last update: 29.06.2017
'-----------------------------------------------------------------------------
    Unload Me
End Sub
