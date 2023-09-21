VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDNameForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "IDNameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IDNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objectTypeId As IdNameFormType
Private curItem As Object


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' инициализация формы
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long

    ' параметры ListView
    ' в виде таблицы
    Me.ListViewList.View = lvwReport
    ' выделяется вся строка
    Me.ListViewList.FullRowSelect = True
    ' запрет изменения значений в самом ListView
    Me.ListViewList.LabelEdit = lvwManual
    ' заголовки столбцов
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To 2
            .add
        Next i
        .Item(1).text = "Код"
        .Item(2).text = "Название"
    End With
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' активация формы - заполнение
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Select Case objectTypeId:
        Case IdNameFormType.edfPlanStatus:
            Me.Caption = "Статусы планирумых работ"
        Case IdNameFormType.edfWallMaterial
            Me.Caption = "Материалы стен"
        Case IdNameFormType.edfWorkType
            Me.Caption = "Типы работ"
    End Select
    Call reloadListView
    Me.TextBoxName.SetFocus
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' обработка выбора элемента в списке
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call initialCurItem(False, Item)
    Me.TextBoxName = curItem.Name
    Me.LabelCurItem.Caption = curItem.Name
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' обработка кнопки закрытия формы
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    If Me.objectTypeId = edfWallMaterial Then
        wallmaterial_list.reload
    ElseIf Me.objectTypeId = edfWorkType Then
        worktype_list.reload
    End If
    Call reloadListView
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' обработка кнопки добавления
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call initialCurItem(True)
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
    If curItem Is Nothing Then Exit Sub
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
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    Dim listX As ListItem
    Dim i As Long, j As Long
    Dim curList As Object
    
    ' заполнение данными
    Select Case objectTypeId:
        Case IdNameFormType.edfPlanStatus:
            Set curList = plan_statuses
        Case IdNameFormType.edfWallMaterial
            Set curList = wallmaterial_list
        Case IdNameFormType.edfWorkType
            Set curList = worktype_list
    End Select
    
    Me.ListViewList.ListItems.Clear
    For i = 1 To curList.count
        Set curItem = curList(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        listX.ListSubItems.add
        listX.ListSubItems(1).text = curItem.Name
        Set curItem = Nothing
    Next i
    Set listX = Nothing
    Set curList = Nothing

    ' ширины столбцов
    Call AppNewAutosizeColumns(Me.ListViewList)
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' добавление/изменение
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If formNotFill Then
        MsgBox "Заполнены не все поля!", vbInformation + vbOKOnly, "Ошибка"
                                                            
        GoTo cleanHandler
    End If
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
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


Private Sub initialCurItem(newFlag As Boolean, _
                                    Optional Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' инициализация текущего элемента
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Select Case objectTypeId
        Case IdNameFormType.edfPlanStatus:
            If newFlag Then
                Set curItem = New plan_status
            Else
                Set curItem = plan_statuses(CStr(Item))
            End If
            
        Case IdNameFormType.edfWallMaterial
            If newFlag Then
                Set curItem = New wallmaterial_class
            Else
                Set curItem = wallmaterial_list(CStr(Item))
            End If
            
        Case IdNameFormType.edfWorkType
            If newFlag Then
                Set curItem = New worktype_class
            Else
                Set curItem = worktype_list(CStr(Item))
            End If
    
    End Select
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' проверка на заполнение полей
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    formNotFill = StrComp(Trim(Me.TextBoxName.Value), "") = 0
End Function


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' очистка всех текстовых полей
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub
