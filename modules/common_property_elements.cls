VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' список видов элементов общего имущества
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' столбцы формы
' 25.05.2022
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fGroup
    fName
    fRequired
    fMax
End Enum


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim CurLItem As common_property_element
    
    With listView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "Код"
            .Item(FormColumns.fGroup + 1).text = "Группа"
            .Item(FormColumns.fName + 1).text = "Название"
            .Item(FormColumns.fRequired + 1).text = "Обязательный"
        End With
        
        ' заполнение данными
        .ListItems.Clear
        For i = 1 To this.count
            Set CurLItem = this(i)
            Set listX = .ListItems.add(, , CurLItem.Id)
            For j = 1 To FormColumns.fMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormColumns.fName).text = CurLItem.Name
            listX.ListSubItems(FormColumns.fGroup).text = CurLItem.Group.Name
            listX.ListSubItems(FormColumns.fRequired).text = BoolToYesNo(CurLItem.IsRequired, trueString:="", falseString:="Нет")
        Next i
    End With
    
    ' ширины столбцов
    Call AppNewAutosizeColumns(listView)
    
    Set CurLItem = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' реализация базового класса - название формы
' 07.12.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "Виды элементов общего имущества дома"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' 18.11.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload(Optional ShowGroupName As Boolean = False)
' ----------------------------------------------------------------------------
' обновление коллекции
' 29.11.2021
' ----------------------------------------------------------------------------
    Call initial(ShowGroupName)
End Sub


Private Sub initial(Optional ShowGroupName As Boolean = False)
' ----------------------------------------------------------------------------
' получение списка
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As common_property_element
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_common_property_element_list"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New common_property_element
        tempItem.add _
            InId:=rst!Id, _
            InGroupId:=rst!group_id, _
            InName:=rst!Name, _
            InRequired:=boolValue(rst!is_required), _
            ShowGroupName:=ShowGroupName
        
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    GoTo cleanHandler

errHandler:
    Dim errMsg As String
    If errorHasNoPrivilegies(Err.Description) Then
        errMsg = "Не хватает прав на получение списка"
    Else
        errMsg = Err.Number & " " & Err.Description
    End If
    MsgBox errMsg, vbExclamation, "Ошибка"

cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set sqlParams = Nothing
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' 18.11.2021
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' 18.11.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As common_property_element
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' 07.12.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
