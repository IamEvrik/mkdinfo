VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "offer_works"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' список работ по предложению к договору
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' столбцы формы
' 22.10.2021
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fYear
    fPriority
    fName
    fSum
    fMax
End Enum
    

Private Sub basicListForm_fillListForm(listView As Object)
    Call Me.fillListForm(listView)
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' реализация базового класса - название формы
' 14.10.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "Предложения по работам"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' 14.10.2021
' ----------------------------------------------------------------------------
    Set this = New Collection
End Sub


Public Sub reload(InBldnId As Long)
' ----------------------------------------------------------------------------
' обновление коллекции
' 14.10.2021
' ----------------------------------------------------------------------------
    Call initial(InBldnId)
End Sub


Private Sub initial(InBldnId As Long)
' ----------------------------------------------------------------------------
' получение списка
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As offer_work
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_work_offers_in_bldn"
    sqlParams.add "InBldnId", InBldnId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New offer_work
        tempItem.add _
            InId:=rst!Id, _
            InName:=rst!work_name, _
            InBldnId:=rst!bldn_id, _
            InYear:=rst!offers_year, _
            InSum:=dblValue(rst!work_sum), _
            InPriority:=rst!Priority
        
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
' 14.10.2021
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
' 14.10.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As offer_work
' ----------------------------------------------------------------------------
' элемент коллекции
' 14.10.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub fillListForm(curListView As listView)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As offer_work
    
    With curListView
        .View = lvwReport       ' в виде таблицы
        .FullRowSelect = True   ' выделяется вся строка
        .LabelEdit = lvwManual  ' запрет изменения значений в самом ListView
        .Gridlines = True       ' сетка
        
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "Код"
            .Item(FormColumns.fName + 1).text = "Название"
            .Item(FormColumns.fPriority + 1).text = "Приоритет"
            .Item(FormColumns.fSum + 1).text = "Сумма"
            .Item(FormColumns.fYear + 1).text = "Год"
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
            listX.ListSubItems(FormColumns.fPriority).text = CurLItem.Priority
            listX.ListSubItems(FormColumns.fSum).text = CurLItem.WorkSum
            listX.ListSubItems(FormColumns.fYear).text = Year(CurLItem.OffersYear)
        Next i
    End With
    
    ' ширины столбцов
    'Call AppNewAutosizeColumns(listView)
    Call AutosizeListViewColumn(curListView, FormColumns.fName)
    curListView.ColumnHeaders(FormColumns.fID + 1).Width = 0
    
    Set CurLItem = Nothing
End Sub


Public Sub ExportList(ByRef InBldn As building_class)
' ----------------------------------------------------------------------------
' создание и печать предложения
' 18.10.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    Dim errMsg As String
    
    On Error GoTo errHandler
    
    sqlString = "create_and_get_work_annex"
    sqlParams.add "InBldnId", InBldn.Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF Or rst.BOF Then GoTo cleanHandler
    
    Dim xml As Object
    
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    If Not xml.LoadXML(rst!outXml) Then
        errMsg = "С сервера получен неверный ответ"
        GoTo errHandler
    End If
    
    Set this = New Collection
    
    Dim xmlChild As Object
    Dim xmlNode As Object
    Dim curItem As offer_work
    
    Set xmlNode = xml.SelectSingleNode("table")
    For Each xmlChild In xmlNode.ChildNodes
        Set curItem = New offer_work
        curItem.add InId:=NOTVALUE, _
        InBldnId:=InBldn.Id, _
        InYear:=NOTDATE, _
        InName:=xmlChild.SelectSingleNode("work_name").text(), _
        InSum:=0, _
        InPriority:=xmlChild.SelectSingleNode("priority").text()
        this.add curItem
        Set curItem = Nothing
    Next xmlChild

    ' создание отчета
    Dim reportWB As Workbook
    Dim reportWS As Worksheet
    Set reportWB = Workbooks.Open(AppConfig.Razdel3FileName)
    Set reportWS = reportWB.Sheets(1)
    Dim firstWorkRow As Integer, rowIdx As Integer
    Dim workCount As Integer
    Dim ppColumn As Integer, workColumn As Integer
    
    With reportWS
        firstWorkRow = .Range("РАБОТА").Row
        rowIdx = firstWorkRow
        workColumn = .Range("РАБОТА").Column
        ppColumn = .Range("ПП").Column
        For Each curItem In this
            workCount = workCount + 1
            .Rows(firstWorkRow).Copy .Rows(rowIdx)
            .Cells(rowIdx, ppColumn) = "3." & workCount & "."
            .Cells(rowIdx, workColumn) = curItem.Name
            rowIdx = rowIdx + 1
        Next curItem
    End With
    
errHandler:
    If Err.Number <> 0 Then
        errMsg = Err.Description
        Err.Clear
    End If
    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation, "Ошибка"
        If Not reportWB Is Nothing Then
            reportWB.Close savechanges:=False
            Application.CutCopyMode = False
        End If
    End If
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set reportWB = Nothing
    Set reportWS = Nothing
    Set rst = Nothing
    Set sqlParams = Nothing
    Set xml = Nothing
End Sub

