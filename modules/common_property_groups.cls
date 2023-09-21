VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' список групп элементов общего имущества
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' столбцы формы
' 21.09.2019
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fName
    fMax
End Enum


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 17.09.2021
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As common_property_group
    
    With listView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "Код"
            .Item(FormColumns.fName + 1).text = "Название"
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
        Next i
    End With
    
    ' ширины столбцов
    Call AppNewAutosizeColumns(listView)
    
    Set CurLItem = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' реализация базового класса - название формы
' 21.09.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "Группы элементов общего имущества дома"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' 21.09.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление коллекции
' 21.09.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' получение списка
' 21.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As common_property_group
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_common_property_group_list"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New common_property_group
        tempItem.add _
            InId:=rst!Id, _
            InName:=rst!Name
        
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
' 21.09.2017
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
' 21.09.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As common_property_group
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' 21.09.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
