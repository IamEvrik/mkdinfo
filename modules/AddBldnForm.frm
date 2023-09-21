VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddBldnForm 
   Caption         =   "Добавить дом"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "AddBldnForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddBldnForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы, установка доступности и заполнение полей
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & ". Сервер " & AppConfig.DBServer
    Call clearForm
    Call reloadComboBox(rcmMd, Me.ComboBoxMO)
    Call reloadComboBox(rcmMC, Me.ComboBoxMC)
    Call reloadComboBox(rcmDogovor, Me.ComboBoxContractType)
End Sub

' ----------------------------------------------------------------------------
' Name: BldnId textbox exit event
' Last update: 15.12.2016
' About: проверка доступности кнопки сохранения при выходе из поля
' ----------------------------------------------------------------------------
Private Sub TextBldnId_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: noBldnId checkbox change event
' Last update: 15.12.2016
' About: изменение галки, есть ли код у дома
' ----------------------------------------------------------------------------
Private Sub CheckBoxNoBldnId_Change()
    Me.TextBldnId.Value = ""
    If Not Me.CheckBoxNoBldnId Then
        Me.TextBldnId.Enabled = True
        Me.TextBldnId.SetFocus
    End If
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: mo combobox change event
' Last update: 06.12.2016
' About: при изменении МО заполняется список населённых пунктов
' ----------------------------------------------------------------------------
Private Sub ComboBoxMO_Change()
    If Me.ComboBoxMO.ListIndex > -1 Then
        Me.ComboBoxVillage.Enabled = True
        Call reloadComboBox(rcmVillage, Me.ComboBoxVillage, initValue:=Me.ComboBoxMO.Value)
    Else
        Me.ComboBoxVillage.Enabled = False
    End If
End Sub

' ----------------------------------------------------------------------------
' Name: village combobox change event
' Last update: 06.12.2016
' About: при изменении села заполняется список улиц
' ----------------------------------------------------------------------------
Private Sub ComboBoxVillage_Change()
    If Me.ComboBoxVillage.ListIndex > -1 Then
        Me.ComboBoxStreet.Enabled = True
        Call reloadComboBox(rcmStreet, Me.ComboBoxStreet, initValue:=Me.ComboBoxVillage.Value)
    Else
        Me.ComboBoxStreet.Enabled = False
    End If
End Sub

' ----------------------------------------------------------------------------
' Name: street combobox change event
' Last update: 15.12.2016
' About: при изменении улицы можно ввести номер дома
' ----------------------------------------------------------------------------
Private Sub ComboBoxStreet_Change()
    If Me.ComboBoxStreet.ListIndex > -1 Then
        Me.TextBldnNo.Enabled = True
    Else
        Me.TextBldnNo.Value = ""
        Me.TextBldnNo.Enabled = False
    End If
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: bldnNo text change event
' Last update: 15.12.2016
' About: при изменении номера дома можно выбрать УК
' ----------------------------------------------------------------------------
Private Sub TextBldnNo_Change()
    If Trim(Me.TextBldnNo.Value) <> "" Then
        Me.ComboBoxMC.Enabled = True
    Else
        Me.ComboBoxMC.Enabled = False
    End If
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: mc combobox change event
' Last update: 06.12.2016
' About: при изменении УК можно выбрать вид договора
' ----------------------------------------------------------------------------
Private Sub ComboBoxMC_Change()
    If Me.ComboBoxMC.ListIndex > -1 Then
        Me.ComboBoxContractType.Enabled = True
    Else
        Me.ComboBoxContractType.Enabled = False
    End If
End Sub

' ----------------------------------------------------------------------------
' Name: contracttype combobox change event
' Last update: 15.12.2016
' About: при изменении вида договора доступна кнопка добавления
' ----------------------------------------------------------------------------
Private Sub ComboBoxContractType_Change()
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: add button click event
' Last update: 13.07.2017
' About: создание объекта
' ----------------------------------------------------------------------------
Private Sub ButtonAdd_Click()
    On Error GoTo errHandler
    
    Dim BldnId As Long
    If Me.CheckBoxNoBldnId.Value Then
        BldnId = NOTVALUE
    Else
        BldnId = CInt(Me.TextBldnId.Value)
    End If
    Dim curBldn As New building_class
    curBldn.create BldnId, Me.ComboBoxStreet.Value, _
                    Me.TextBldnNo.Value, Me.ComboBoxMC.Value, _
                    Me.ComboBoxContractType.Value
    Set curBldn = Nothing
    MsgBox "Адрес успешно добавлен"
    Call cleanform
    GoTo cleanHandler

errHandler:
    MsgBox "Не удалось создать дом" & vbCrLf & Err.Description, _
                                vbCritical, "Ошибка"

cleanHandler:
End Sub

' ----------------------------------------------------------------------------
' Name: cancel button click event
' Last update: 06.12.2016
' About: отмена, закрытие формы
' ----------------------------------------------------------------------------
Private Sub ButtonCancel_Click()
    Unload Me
End Sub

' ----------------------------------------------------------------------------
' Name: clearForm method
' Last update: 15.12.2016
' About: очистка полей формы
' ----------------------------------------------------------------------------
Private Sub clearForm()
    ' очистка значение
    Me.ComboBoxContractType.Clear
    Me.ComboBoxMC.Clear
    Me.ComboBoxMO.Clear
    Me.ComboBoxStreet.Clear
    Me.ComboBoxVillage.Clear
    Me.TextBldnNo.Value = ""
    Me.TextBldnId.Value = ""
    Me.CheckBoxNoBldnId.Value = True
    ' доступны только нужные элементы
    Me.ComboBoxContractType.Enabled = False
    Me.ComboBoxMC.Enabled = False
    Me.ComboBoxStreet.Enabled = False
    Me.ComboBoxVillage.Enabled = False
    Me.TextBldnNo.Enabled = False
    Me.TextBldnId.Enabled = False
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: cleanForm method
' Last update: 15.12.2016
' About: очистка полей формы после сохранения
' ----------------------------------------------------------------------------
Private Sub cleanform()
    Me.TextBldnId.Value = ""
    Me.TextBldnNo.Value = ""
    Call buttonVisible
End Sub

' ----------------------------------------------------------------------------
' Name: buttonVisible method
' Last update: 15.12.2016
' About: возвращение доступности элементов к начальному состоянию
' ----------------------------------------------------------------------------
Private Sub buttonVisible()
    If Me.ComboBoxContractType.ListIndex > -1 And _
                    Me.ComboBoxMC.ListIndex > -1 And _
                    Me.ComboBoxMO.ListIndex > -1 And _
                    Me.ComboBoxStreet.ListIndex > -1 And _
                    Me.ComboBoxVillage.ListIndex > -1 And _
                    (Me.TextBldnId.Value <> "" Or Me.CheckBoxNoBldnId) And _
                    Me.TextBldnNo.Value <> "" Then
        Me.ButtonAdd.Enabled = True
    Else
        Me.ButtonAdd.Enabled = False
    End If
End Sub
