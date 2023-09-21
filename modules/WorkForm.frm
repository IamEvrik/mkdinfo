VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorkForm 
   Caption         =   "Ввод новой работы"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   OleObjectBlob   =   "WorkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' Переменные формы
' ----------------------------------------------------------------------------
Public BldnId As Long               ' код дома
Public changedWork As work_class    ' изменяемая работа (если изменение)
Public planWork As plan_work_class  ' выполнение плановой работы
Public formContractorId As Long     ' подрядчик
Public formPrefGWT As Long          ' вид ремонта
Public formMcId As Long             ' ук


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' инициализация формы, заполнение данных
' Last update: 29.09.2020
' ----------------------------------------------------------------------------
    ' инициализация переменных
    Call reloadComboBox(rcmContractor, Me.ComboBoxContractor)
    Call reloadComboBox(rcmGWT, Me.ComboBoxGlobalWorkType)
    Call reloadComboBox(rcmWorkType, Me.ComboBoxWorkType)
    Call reloadComboBox(rcmTermDESC, Me.ComboBoxTerms)
    Call reloadComboBox(rcmMC, Me.ComboBoxMC, defValue:=formMcId)
    Call reloadComboBox(rcmFSources, Me.ComboBoxFSource)

    ' какую дату выставлять. если есть сохранённая, то её, иначе сегодня
    Dim tmpDateTerm As Long
    tmpDateTerm = getUseWorkDate()
    If tmpDateTerm <> NOTVALUE Then
        Call selectComboBoxValue(Me.ComboBoxTerms, tmpDateTerm)
    Else
        Me.ComboBoxTerms.ListIndex = 0
    End If
        
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы, заполнение полей, если изменение
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    If formPrefGWT = SERVICE_GLOBAL_TYPE Then
        Call selectComboBoxValue(Me.ComboBoxContractor, formContractorId)
        Me.ComboBoxContractor.Enabled = False
        Me.ComboBoxGlobalWorkType.Enabled = False
        Call selectComboBoxValue(Me.ComboBoxFSource, 0)
        Me.ComboBoxFSource.Enabled = False
    Else
        Call selectComboBoxValue(Me.ComboBoxFSource, 1)
    End If
    If Not Me.changedWork Is Nothing Then
        Call selectComboBoxValue(Me.ComboBoxTerms, changedWork.WorkDate)
        Call selectComboBoxValue(Me.ComboBoxGlobalWorkType, _
                                            changedWork.GWT.Id)
        Call selectComboBoxValue(Me.ComboBoxWorkType, _
                                            changedWork.WorkKind.workType.Id)
        Call selectComboBoxValue(Me.ComboBoxWorkKind, _
                                            changedWork.WorkKind.Id)
        Call selectComboBoxValue(Me.ComboBoxContractor, _
                                            changedWork.Contractor.Id)
        Call selectComboBoxValue(Me.ComboBoxMC, changedWork.MC.Id)
        Call selectComboBoxValue(Me.ComboBoxFSource, _
                                            changedWork.financeSource.Id)
        Me.TextBoxNote = Me.changedWork.Note
        Me.TextBoxSum = CDbl(Me.changedWork.sum)
        Me.TextBoxVolume = Me.changedWork.Volume
        Me.TextBoxDogovor = Me.changedWork.Dogovor
        Me.TextBoxSi = Me.changedWork.Si
        Me.CheckBoxNotPrint = Not Me.changedWork.PrintFlag
        Me.TextBoxPrivateNote = Me.changedWork.privateNote
    ElseIf Not planWork Is Nothing Then
        Call selectComboBoxValue(Me.ComboBoxGlobalWorkType, _
                                            planWork.GWT.Id)
        Call selectComboBoxValue(Me.ComboBoxWorkType, _
                                            planWork.WorkKind.workType.Id)
        Call selectComboBoxValue(Me.ComboBoxWorkKind, _
                                            planWork.WorkKind.Id)
        Call selectComboBoxValue(Me.ComboBoxContractor, _
                                            planWork.Contractor.Id)
        Call selectComboBoxValue(Me.ComboBoxMC, planWork.MC.Id)
        Me.TextBoxNote = planWork.Note
        Me.TextBoxSum = CDbl(planWork.sum)
        Me.BldnId = planWork.BldnId
        Me.ComboBoxTerms.ListIndex = -1
    Else
        Call selectComboBoxValue(Me.ComboBoxGlobalWorkType, formPrefGWT)
        Call selectComboBoxValue(Me.ComboBoxMC, formMcId)
    End If
        
End Sub


Private Sub ComboBoxWorkType_Change()
' ----------------------------------------------------------------------------
' при изменении типа работ заполняются виды работы.
' Last update: 22.04.2016
' ----------------------------------------------------------------------------
    If Me.ComboBoxWorkType.ListIndex > -1 Then
        Call reloadComboBox(rcmWorkKind, Me.ComboBoxWorkKind, _
                                        initValue:=Me.ComboBoxWorkType.Value)
    End If
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки "Сохранить"
' Last update: 26.02.2019
' ----------------------------------------------------------------------------
    Dim strSum As String
    Dim workSum As Currency
    
    If Me.ComboBoxWorkKind.ListIndex = -1 Then
    ' проверка на заполнение работы
        Call setMsg("Укажите вид работы")
        Exit Sub
    End If
    
    If Trim(Me.TextBoxSum.Value) = "" Then
    ' проверка введенной суммы
        Call setMsg("Введите сумму")
        Exit Sub
    Else
    ' преобразование суммы в правильный вид (иначе сообщение об ошибке)
        workSum = dblValue(Trim(Me.TextBoxSum.Value))
        If workSum = NOTVALUE Then
            setMsg ("Ошибка в сумме работы")
            Exit Sub
        End If
    End If
    
    If Me.ComboBoxFSource.ListIndex = -1 Then
        setMsg ("Выберите источник финансирования")
        Exit Sub
    End If
    
    If Me.ComboBoxTerms.ListIndex = -1 Then
        setMsg ("Выберите дату")
        Exit Sub
    End If
    
    ' проверка на то, что УК имеет признак УК
    Dim notManageMark As Boolean
    Dim Answer As Integer
    Dim tmpMC As New uk_class
    tmpMC.initial Me.ComboBoxMC.Value
    notManageMark = tmpMC.notManage
    Set tmpMC = Nothing
    If notManageMark Then
        Answer = MsgBox("Правильно ли указана УК!", _
                                vbYesNo + vbExclamation, "Проверьте выбор")
        If Answer = vbNo Then Exit Sub
    End If
    
    ' сохранение работы
    On Error GoTo errHandler
    If Not Me.changedWork Is Nothing Then
        Me.changedWork.update newGWT:=Me.ComboBoxGlobalWorkType.Value, _
                        newWK:=Me.ComboBoxWorkKind.Value, _
                        newDate:=Me.ComboBoxTerms.Value, _
                        newSum:=workSum, _
                        newSi:=Me.TextBoxSi.Value, _
                        newVolume:=Me.TextBoxVolume.Value, _
                        newNote:=Me.TextBoxNote.Value, _
                        newContractor:=Me.ComboBoxContractor.Value, _
                        newMC:=Me.ComboBoxMC.Value, _
                        newDogovor:=Me.TextBoxDogovor.Value, _
                        newFSource:=Me.ComboBoxFSource.Value, _
                        newPF:=Not Me.CheckBoxNotPrint.Value, _
                        newPrivateNote:=Me.TextBoxPrivateNote.Value
    Else
        Dim curWork As New work_class
    
        curWork.create BldnId:=Me.BldnId, workSum:=workSum, _
                        WorkDate:=Me.ComboBoxTerms.Value, _
                        workKindID:=Me.ComboBoxWorkKind.Value, _
                        Si:=Me.TextBoxSi.Value, _
                        workVolume:=Me.TextBoxVolume.Value, _
                        workNote:=Me.TextBoxNote.Value, _
                        contractorId:=Me.ComboBoxContractor.Value, _
                        mcId:=Me.ComboBoxMC.Value, _
                        Dogovor:=Me.TextBoxDogovor.Value, _
                        gwtId:=Me.ComboBoxGlobalWorkType.Value, _
                        PrintFlag:=(Not Me.CheckBoxNotPrint.Value), _
                        financeSource:=Me.ComboBoxFSource.Value, _
                        privateNote:=Me.TextBoxPrivateNote
        Call saveUseWorkDate(curWork.WorkDate)
        If Not planWork Is Nothing Then
            planWork.setDone workRef:=curWork.Id
        End If
    End If
    
    BuildingForm.workChanged = True
    BuildingForm.planWorkChanged = True
    Unload Me
    BuildingForm.Show
    GoTo cleanHandler

errHandler:
    setMsg (Err.Description)
    GoTo cleanHandler
    
cleanHandler:
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки "отмена"
' Last update: 07.06.2016
' ----------------------------------------------------------------------------
    Unload Me           ' выгружается форма
    BuildingForm.workChanged = False
    BuildingForm.Show   ' показ формы здания
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' запрет закрытия формы крестиком,
'                    т.к. после этого некорректно работает показ формы МКД
' Last update: 01.03.2016
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' уничтожение формы
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    Set changedWork = Nothing
    Set planWork = Nothing
End Sub


Private Sub setMsg(msgText As String, Optional isError As Boolean = True)
' ----------------------------------------------------------------------------
' вывод сообщения
' Last update: 19.09.2018
' ----------------------------------------------------------------------------
    Me.LabelMsg.Caption = msgText
    Me.LabelMsg.ForeColor = IIf(isError, RGB(255, 0, 0), RGB(0, 0, 0))
End Sub
