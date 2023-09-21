VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GlobalWorkForm 
   Caption         =   "Виды ремонтов"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "GlobalWorkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GlobalWorkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New globalWorkType_class


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    globalWorkType_list.reload
    Call reloadListView
End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' активация формы, заполнение полей
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    ' в виде таблицы
    Me.ListViewList.View = lvwReport
    ' выделяется вся строка
    Me.ListViewList.FullRowSelect = True
    ' запрет изменения значений в самом ListView
    Me.ListViewList.LabelEdit = lvwManual
    ' заголовки столбцов
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To FormGWTEnum.fgwtMax
            .add
        Next i
        .Item(FormGWTEnum.fgwtId + 1).text = "Код"
        .Item(FormGWTEnum.fgwtName + 1).text = "Название"
        .Item(FormGWTEnum.fgwtNote + 1).text = "Примечание"
    End With
    
    Call reloadListView
    
    Me.TextBoxName.SetFocus
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' обработка выбора элемента в списке
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = globalWorkType_list(CStr(Item))
    Me.TextBoxName = curItem.Name
    Me.TextBoxNote = curItem.Note
    Me.LabelCurItem.Caption = curItem.Name
End Sub


Private Sub ListViewList_ColumnClick( _
                                ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' сортировка при щелчке на столбце
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' обработка кнопки закрытия формы
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' обработка кнопки добавления
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New globalWorkType_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub BtnChange_Click()
' ----------------------------------------------------------------------------
' обработка кнопки изменения
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' обработка кнопки очистки выбора
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub BtnDelete_Click()
' ----------------------------------------------------------------------------
' обработка кнопки удаления
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        ' запрос подтверждения
        If Not ConfirmDeletion(curItem.Name) Then Exit Sub
        
        On Error GoTo errHandler
        curItem.delete
        
        ' перерисовка
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler
        
errHandler:
        If Err.Number = ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "Ошибка удаления"
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
cleanHandler:
    End If
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' заполнение данными ListView
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim listX As ListItem
    Dim i As Long, j As Long
    
    ' заполнение данными
    Me.ListViewList.ListItems.Clear
    For i = 1 To globalWorkType_list.count
        Set curItem = globalWorkType_list(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        For j = 1 To FormGWTEnum.fgwtMax - 1
            listX.ListSubItems.add
        Next j
        listX.ListSubItems(FormGWTEnum.fgwtName).text = curItem.Name
        listX.ListSubItems(FormGWTEnum.fgwtNote).text = curItem.Note
        Set curItem = Nothing
    Next i
    Set listX = Nothing

    ' ширины столбцов
    Call AppNewAutosizeColumns(Me.ListViewList)
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' добавление/изменения вида ремонта
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If formNotFill Then
        MsgBox "Заполнены не все поля!", vbInformation + vbOKOnly, "Ошибка"
                                                            
        GoTo cleanHandler
    End If
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
                        newNote:=Me.TextBoxNote.Value, _
                        addNew:=addFlag
        
        ' перерисовка формы
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler
        
errHandler:
        If Err.Number = ERROR_NOT_UNIQUE Then
            MsgBox Err.Description, vbInformation, "Ошибка"
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
        
cleanHandler:
    End If
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' проверка на заполнение полей
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    formNotFill = (StrComp(Trim(Me.TextBoxName.Value), "") = 0)
End Function


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' очистка всех текстовых полей
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.TextBoxNote.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub
