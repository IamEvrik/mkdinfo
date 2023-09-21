VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UKForm 
   Caption         =   "Справочник управляющих компаний"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "UKForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UKForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New uk_class


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    uk_list.reload
    Call reloadListView
End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' активация формы, заполнение полей
' Last update: 17.08.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    Me.Caption = Me.Caption & " " & AppConfig.DBServer
    
    ' в виде таблицы
    Me.ListViewList.View = lvwReport
    ' выделяется вся строка
    Me.ListViewList.FullRowSelect = True
    ' запрет изменения значений в самом ListView
    Me.ListViewList.LabelEdit = lvwManual
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To FormMCEnum.fmcMax
            .add
        Next i
        .Item(FormMCEnum.fmcID + 1).text = "Код"
        .Item(FormMCEnum.fmcName + 1).text = "Название"
        .Item(FormMCEnum.fmcReportName + 1).text = "Сокр. название"
        .Item(FormMCEnum.fmcNotManage + 1).text = "Управляющая"
        .Item(FormMCEnum.fmcChiefEngineer + 1).text = "Главный инженер"
        .Item(FormMCEnum.fmcDirector + 1).text = "Директор"
    End With
    Me.ListViewList.Sorted = True
    Me.ListViewList.SortKey = FormMCEnum.fmcID
    
    Call reloadListView
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы - очистка полей
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' обработка выбора управляющей компании в списке
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set curItem = uk_list(CStr(Item))
    ' заполнение полей
    Me.TextBoxName = curItem.Name
    Me.TextBoxReportName = curItem.reportName
    Me.CheckBoxNotManage = curItem.notManage
    Me.LabelCurItem.Caption = curItem.Name
End Sub


Private Sub ListViewList_ColumnClick(ByVal ColumnHeader As _
                                                    MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' сортировка при щелчке на столбце
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' обработка кнопки закрытия формы
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' обработка кнопки очистки формы
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' обработка кнопки добавления УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New uk_class
    Call process(addFlag:=True)
End Sub


Private Sub BtnChange_Click()
' ----------------------------------------------------------------------------
' обработка кнопки изменения УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub BtnDelete_Click()
' ----------------------------------------------------------------------------
' обработка кнопки удаления УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        ' подтверждение удаления
        If Not ConfirmDeletion(curItem.Name) Then Exit Sub
        
        On Error GoTo errHandler
        curItem.delete
        
        Call reloadListView
        Call clearTextBox
        
errHandler:
        If Err.Number = ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "Ошибка"
        Else
            MsgBox Err.Number & "-->" & Err.Description, vbCritical
        End If
    End If
End Sub


Private Sub BtnShowEmployee_Click()
' ----------------------------------------------------------------------------
' обработки нажатия кнопки "Сотрудники", показ формы сотрудников
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        Me.Hide
        Call RunEmployeeForm(curItem)
    End If
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' заполнение данными ListView
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As ListItem
    
    ' заполнение данными
    Me.ListViewList.ListItems.Clear
    For i = 1 To uk_list.count
        Set curItem = uk_list(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormMCEnum.fmcMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormMCEnum.fmcName).text = curItem.Name
        listX.ListSubItems(FormMCEnum.fmcReportName).text = curItem.reportName
        listX.ListSubItems(FormMCEnum.fmcDirector).text = curItem.Director.FIO
        listX.ListSubItems(FormMCEnum.fmcChiefEngineer).text = _
                                                    curItem.ChiefEngineer.FIO
        listX.ListSubItems(FormMCEnum.fmcNotManage).text = _
                                            BoolToYesNo(Not curItem.notManage)
    Next i
    Set curItem = Nothing
    
    ' автоматическая ширина столбцов
    Call AppNewAutosizeColumns(Me.ListViewList)
    
    Set listX = Nothing
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' добавление/изменение в базе
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Or addFlag Then
        If formNotFill Then
            MsgBox "Заполнены не все необходимые поля", vbInformation, _
                                                                    "Ошибка"
            Exit Sub
        End If
            
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
                        newReport:=Me.TextBoxReportName.Value, _
                        newNotManage:=Me.CheckBoxNotManage.Value, _
                        addNew:=addFlag
                
        Call reloadListView
        Call clearTextBox
    End If
    GoTo cleanHandler

errHandler:
    If Err.Number = ERROR_NOT_UNIQUE Then
        MsgBox Err.Description, vbInformation, "Ошибка"
    Else
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

cleanHandler:
End Sub


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' очистка всех текстовых полей
' Last update: 04.03.2016
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.TextBoxReportName.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
    Me.CheckBoxNotManage.Value = False
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' проверка на заполнение полей
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    formNotFill = False
    If StrComp(Trim(Me.TextBoxName.Value), "") = 0 Or _
                    StrComp(Trim(Me.TextBoxReportName.Value), "") = 0 Then
        formNotFill = False
    End If
End Function
