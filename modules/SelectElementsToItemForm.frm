VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectElementsToItemForm 
   Caption         =   "”становка прав доступа дл€ пользователей"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   OleObjectBlob   =   "SelectElementsToItemForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectElementsToItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_Caption As String
Private m_EnableItemSelect As Boolean
Private m_FormType As SelectElementsFormTypeEnum
Private m_ItemId As Long


' ----------------------------------------------------------------------------
' установка параметров формы
' 26.11.2021
' ----------------------------------------------------------------------------
Property Let SetCaption(newCaption As String)
    m_Caption = newCaption
End Property
Property Let SetEnableItemSelect(newState As Boolean)
    m_EnableItemSelect = newState
End Property
Property Let SetFormType(newType As SelectElementsFormTypeEnum)
    m_FormType = newType
End Property
Property Let SetItemId(newId As Long)
    m_ItemId = newId
End Property

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' »нициализаци€ формы, начальна€ установка свойств
' 26.11.2021
' ----------------------------------------------------------------------------
    m_Caption = ""
    m_EnableItemSelect = True
    m_FormType = seftBldnCpe
    m_ItemId = NOTVALUE
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' јктиваци€ формы
' 26.11.2021
' ----------------------------------------------------------------------------
    Me.Caption = m_Caption & ". —ервер " & AppConfig.DBServer
    Call reloadComboItems
    Me.ComboBoxItem.Enabled = m_EnableItemSelect
End Sub


Private Sub ComboBoxItem_Change()
' ----------------------------------------------------------------------------
' ѕри выборе элемента заполнение списков
' 25.11.2021
' ----------------------------------------------------------------------------
    If Me.ComboBoxItem.ListIndex > -1 Then
        Call reloadLists
    End If
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' перенос элемента в добавленные
' 25.11.2021
' ----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxAvailable, Me.ListBoxSelected)
End Sub


Private Sub ButtonDelete_Click()
' ----------------------------------------------------------------------------
' удаление элемента из списка добавленных
' 25.11.2021
' ----------------------------------------------------------------------------
    Call MoveListBoxElements(Me.ListBoxSelected, Me.ListBoxAvailable)
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' запись изменений, обновление списков
' 26.11.2021
' ----------------------------------------------------------------------------
    Dim sqlString As String, sqlParams As New Dictionary
    Dim param As String
    
    Call setMsg("")
    
    sqlString = "add_common_property_elements_to_bldn_list"
    If Me.ListBoxSelected.ListCount > 0 Then
        On Error GoTo errHandler
        Dim selectedItems As New Collection
        Dim idx As Integer
        For idx = 1 To Me.ListBoxSelected.ListCount
            selectedItems.add Me.ListBoxSelected.list(idx - 1, ComboColumns.ccId)
        Next idx
        param = "{" & Join(CollectionToArray(selectedItems), ",") & "}"
        sqlParams.add "InBldnList", "{" & Me.ComboBoxItem & "}"
        sqlParams.add "InElementId", param
        DBConnection.RunQuery sqlString, sqlParams
        Call reloadLists
    End If
    GoTo cleanHandler
    
errHandler:
    Call setMsg(Err.Description, True)
    
cleanHandler:
    If Not m_EnableItemSelect Then Unload Me
End Sub


Private Sub reloadComboItems()
' ----------------------------------------------------------------------------
' заполнение ComboBox
' 26.11.2021
' ----------------------------------------------------------------------------
    If m_FormType = SelectElementsFormTypeEnum.seftBldnCpe Then
        Call reloadComboBox(rcmListManagedBldnAddressIdByMD, Me.ComboBoxItem, _
                initValue:=ALLVALUES, defValue:=m_ItemId)
        If m_ItemId = NOTVALUE Then Me.ComboBoxItem.ListIndex = 0
    End If
End Sub


Private Sub reloadLists()
' ----------------------------------------------------------------------------
' обновление списков
' 26.11.2021
' ----------------------------------------------------------------------------
    If m_FormType = SelectElementsFormTypeEnum.seftBldnCpe Then
        Call reloadComboBox(rcmCommonPropertyElementsNotInBldn, _
                Me.ListBoxAvailable, Me.ComboBoxItem.Value)
    End If
    Me.ListBoxSelected.Clear
End Sub


Private Sub setMsg(msgText As String, Optional isError = False)
' ----------------------------------------------------------------------------
' вывод сообщени€
' 25.11.2021
' ----------------------------------------------------------------------------
    Me.LabelInfo.Caption = msgText
    Me.LabelInfo.ForeColor = IIf(isError, RGB(255, 0, 0), RGB(0, 0, 0))
End Sub
