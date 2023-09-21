VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StreetForm 
   Caption         =   "Список улиц"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13755
   OleObjectBlob   =   "StreetForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StreetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curItem As New address_street_class


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
        For i = 1 To FormStreetEnum.fsMax
            .add
        Next i
        .Item(FormStreetEnum.fsId + 1).text = "Код"
        .Item(FormStreetEnum.fsName + 1).text = "Название"
        .Item(FormStreetEnum.fsSite + 1).text = "Для сайта"
        .Item(FormStreetEnum.fsVillage + 1).text = "Населённый пункт"
    End With
    
    Call reloadListView
    Call reloadComboBox(rcmStreetTypes, Me.ComboBoxStreetType)
    Call reloadComboBox(rcmVillage, Me.ComboBoxVillage)
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' закрытие формы
' Last update: 21.03.2016
' ----------------------------------------------------------------------------
    Set curItem = Nothing
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' при выборе пункта - заполнение полей и внутренней переменной
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    curItem.initial CLng(Item)
    Me.TextBoxName = curItem.StreetName
    Me.TextBoxSite = curItem.Site
    Call selectComboBoxValue(Me.ComboBoxStreetType, curItem.StreetType.Id)
    Call selectComboBoxValue(Me.ComboBoxVillage, curItem.Village.Id)
    Me.CheckBoxHasNoName.Value = (curItem.Name = NOTSTRING)
    Me.LabelCurItem.Caption = curItem.FullName & "(" & _
                                                    curItem.Village.Name & ")"
End Sub


Private Sub ListViewList_ColumnClick( _
                            ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' ----------------------------------------------------------------------------
' сортировка при щелчке на столбце
' Last update: 21.03.2016
' ----------------------------------------------------------------------------
    Me.ListViewList.SortKey = ColumnHeader.Index - 1
End Sub


Private Sub CheckBoxHasNoName_Click()
' ----------------------------------------------------------------------------
' включение/выключение доступности поля ввода названия улицы
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Me.TextBoxName.Enabled = Not Me.CheckBoxHasNoName.Value
    Me.ComboBoxStreetType.Enabled = Not Me.CheckBoxHasNoName.Value
    If Me.CheckBoxHasNoName.Value Then Me.ComboBoxStreetType.ListIndex = 0
    Me.TextBoxName.Value = ""
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' обработка кнопки закрытия формы
' Last update: 21.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' обработка кнопки очистки формы
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' обработка кнопки добавления
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Set curItem = New address_street_class
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub BtnChange_Click()
' ----------------------------------------------------------------------------
' обработка кнопки изменения
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub BtnDelete_Click()
' ----------------------------------------------------------------------------
' обработка кнопки удаления
' Last update: 22.06.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Then
        ' подтверждение удаления
        If Not ConfirmDeletion(curItem.Village.Name & " " & curItem.Name) _
                                                                Then Exit Sub
        
        On Error GoTo errHandler
        curItem.delete
        
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler:
        
errHandler:
        If ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "Ошибка"
        Else
            MsgBox Err.Number & "-->" & Err.Description, vbCritical
        End If
cleanHandler:
    End If
End Sub


Private Sub reloadListView()
' -----------------------------------------------------------------------------
' заполнение данными ListView
' Last update: 24.03.2018
' -----------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim listX As ListItem
    Dim contList As address_street_list
    
    ' заполнение данными
    Set contList = New address_street_list
    With Me.ListViewList.ListItems
        .Clear
        For i = 1 To contList.count
            Set listX = .add(, , contList(i).Id)
            For j = 1 To FormStreetEnum.fsMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormStreetEnum.fsName).text = contList(i).FullName
            listX.ListSubItems(FormStreetEnum.fsSite).text = contList(i).Site
            listX.ListSubItems(FormStreetEnum.fsVillage).text = _
                                                    contList(i).Village.Name
        Next i
    End With
    
    Set contList = Nothing
    Set listX = Nothing

    ' ширины столбцов
    Call AppNewAutosizeColumns(Me.ListViewList)

End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' добавление/изменение в базе
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    If curItem.Id <> NOTVALUE Or addFlag Then
        If formNotFill Then
            MsgBox "Заполнены не все необходимые поля", vbInformation, _
                                                                    "Ошибка"
            Exit Sub
        End If
            
        On Error GoTo errHandler
        curItem.update newName:=IIf(Me.CheckBoxHasNoName, NOTSTRING, _
                                                Trim(Me.TextBoxName.Value)), _
                        newVillage:=Me.ComboBoxVillage.Value, _
                        newSite:=Me.TextBoxSite.Value, _
                        newStreetType:=Me.ComboBoxStreetType.Value, _
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
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.TextBoxSite.Value = ""
    Set curItem = Nothing
    Me.LabelCurItem.Caption = ""
    Me.ComboBoxVillage.ListIndex = -1
    Me.ComboBoxStreetType.ListIndex = -1
    Me.CheckBoxHasNoName.Value = False
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' проверка на заполнение полей
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    formNotFill = False
    If Not Me.CheckBoxHasNoName.Value Then
        If Trim(Me.TextBoxName.Value) = "" Or _
                                    Me.ComboBoxVillage.ListIndex = -1 Or _
                                    Me.ComboBoxStreetType.ListIndex = -1 Then
            formNotFill = True
        End If
    End If
End Function
