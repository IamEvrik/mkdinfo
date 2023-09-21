VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "expense_items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' список статей расходов (статический)
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' столбцы формы
' 25.10.2021
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fShortName
    fService
    fGroupName
    fName1
    fName2
    fGisGuid
    fPriority
    fAsGroupName
    fMax
End Enum


Private Sub basicListForm_fillListForm(ByRef curListView As Object)
'-----------------------------------------------------------------------------
' заполнение статей расходов
' 25.10.2021
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As expense_item
    
    
    With curListView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "Код"
            .Item(FormColumns.fGisGuid + 1).text = "GUID ГИС"
            .Item(FormColumns.fName1 + 1).text = "Название 1"
            .Item(FormColumns.fName2 + 1).text = "Название 2"
            .Item(FormColumns.fShortName + 1).text = "Краткое название"
            .Item(FormColumns.fService + 1).text = "Услуга"
            .Item(FormColumns.fGroupName + 1).text = "Группа"
            .Item(FormColumns.fPriority + 1).text = "Приоритет"
            .Item(FormColumns.fAsGroupName + 1).text = "Как название"
        End With
        
        ' заполнение данными
        .ListItems.Clear
        For i = 1 To this.count
            Set CurLItem = this(i)
            Set listX = .ListItems.add(, , CurLItem.Id)
            For j = 1 To FormColumns.fMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormColumns.fGisGuid).text = CurLItem.GisGuid
            listX.ListSubItems(FormColumns.fName2).text = CurLItem.Name2
            listX.ListSubItems(FormColumns.fName1).text = CurLItem.Name1
            listX.ListSubItems(FormColumns.fShortName).text = _
                                                        CurLItem.ShortName
            listX.ListSubItems(FormColumns.fService).text = _
                                                    CurLItem.UkService.Name
            listX.ListSubItems(FormColumns.fAsGroupName).text = _
                                        BoolToYesNo(CurLItem.UseAsGroupName)
            listX.ListSubItems(FormColumns.fGroupName).text = _
                                                        CurLItem.Group.Name
            listX.ListSubItems(FormColumns.fPriority).text = _
                                                    CurLItem.ReportPriority
        Next i
    End With
    
    ' ширины столбцов
    Call AppNewAutosizeColumns(curListView)
    
    Set CurLItem = Nothing
    Set listX = Nothing
End Sub

Private Property Get basicListForm_Title() As String
    basicListForm_Title = "Статьи расходов"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' Last update: 21.06.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление коллекции
' Last update: 21.06.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' получение списка
' 25.10.2021
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As expense_item
    
    Set this = New Collection
    
    Set rst = DBConnection.GetQueryRecordset("get_expense_items", New Dictionary)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New expense_item
        tempItem.add _
            newId:=rst!Id, _
            newName1:=DBgetString(rst!Name1), _
            newName2:=DBgetString(rst!Name2), _
            newShortName:=DBgetString(rst!short_name), _
            newGisGuid:=DBgetString(rst!gis_guid), _
            newUkServiceId:=rst!uk_service_id, _
            newGroupId:=longValue(rst!group_id), _
            newReportPriority:=longValue(rst!report_priority), _
            newUseAsGroupName:=boolValue(rst!use_as_group_name)
        
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 21.06.2018
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
' Last update: 21.06.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As expense_item
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 21.06.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property

