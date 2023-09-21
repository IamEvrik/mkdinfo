VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BldnReportSelectForm 
   Caption         =   "Выберите дома"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "BldnReportSelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BldnReportSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fReportType As BldnListReportType             ' вид отчета
Private isUpdate As Boolean

Property Let reportType(rType As BldnListReportType)
    fReportType = rType
End Property


Property Get reportType() As BldnListReportType
    reportType = fReportType
End Property


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы, заполнение списков
' 27.09.2022
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & ". Сервер " & AppConfig.DBServer
    If Me.reportType = blrWorkCompletition Then
        Call paintWorkCompletition
    Else
        Call reloadComboBox(rcmListBldnAddressId, Me.ListBoxAvailable)
        Call reloadWorksYearCombo
    End If
End Sub


Private Sub CBOneSelect_Click()
' ----------------------------------------------------------------------------
' выбор одного/нескольких пунктов
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxAvailable, Me.ListBoxSelected)
End Sub


Private Sub CBOneUnselect_Click()
' ----------------------------------------------------------------------------
' отмена выбора одного/нескольких пунктов
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxSelected, Me.ListBoxAvailable)
End Sub


Private Sub CBAllSelect_Click()
' ----------------------------------------------------------------------------
' выбор всех пунктов
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxAvailable, Me.ListBoxSelected, True)
End Sub


Private Sub CBAllUnselect_Click()
' ----------------------------------------------------------------------------
' отмена выбора всех пунктов
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxSelected, Me.ListBoxAvailable, True)
End Sub


Private Sub CBCreateReport_Click()
' ----------------------------------------------------------------------------
' формирование отчётов
' 28.09.2022
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To Me.ListBoxSelected.ListCount - 1
        Me.LabelStatus.Caption = "Формируется отчёт " & _
                                            Me.ListBoxSelected.list(i, ccname)
        DoEvents
        
        If Me.reportType = blrWorkCompletition Then
            Call ReportBldnWorkCompletition(Me.ListBoxSelected.list(i, ccId), _
                    TermId:=Me.ComboBoxYearList.Value)
        Else
            Call BldnPassport(ItemId:=Me.ListBoxSelected.list(i, ccId), _
                                not_show_sum:=Me.CheckBoxNotSum, _
                                year_report:=Me.ComboBoxYearList.Value)
        End If
    Next i
    Me.Hide
    Unload Me
End Sub


Private Sub paintWorkCompletition()
' ----------------------------------------------------------------------------
' отрисовка формы для акта приёмки работ
' 27.09.2022
' ----------------------------------------------------------------------------
    Me.CheckBoxNotSum.Visible = False
    Me.Label2.Caption = "Вид договора"
    isUpdate = False
    Call reloadComboBox(rcmBldnExpenseTerms, Me.ComboBoxYearList, initValue:=ALLVALUES)
    Call reloadComboBox(rcmDogovor, Me.ComboBox2, addAllItems:=True)
    Call reloadComboBox(rcmListBldnAddressIdByMD, Me.ListBoxAvailable, initValue:=ALLVALUES, initValue2:=Me.ComboBox2.Value)
    isUpdate = True
End Sub


Private Sub ComboBox2_Change()
' ----------------------------------------------------------------------------
' обновление списка домов
' 27.09.2022
' ----------------------------------------------------------------------------
    If isUpdate Then
        Call reloadComboBox(rcmListBldnAddressIdByMD, _
                Me.ListBoxAvailable, _
                initValue:=ALLVALUES, _
                initValue2:=Me.ComboBox2.Value)
    End If
End Sub


Private Sub reloadWorksYearCombo()
' ----------------------------------------------------------------------------
' заполнение списка годов
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    Dim tmpCol As New Collection
    Dim i As Long
    
    Set tmpCol = worksYears(gwtId:=NOTVALUE, BldnId:=ALLVALUES)
    Me.ComboBoxYearList.Clear
    For i = 1 To tmpCol.count
        Me.ComboBoxYearList.AddItem tmpCol(i)
    Next i
    
    If Me.ComboBoxYearList.ListCount > -1 Then
        Me.ComboBoxYearList.ListIndex = 0
    End If
End Sub
