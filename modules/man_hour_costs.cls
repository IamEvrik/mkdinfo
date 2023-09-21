VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "man_hour_costs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' стоимости человекочаса
' ----------------------------------------------------------------------------

Private m_this As Collection


Private Enum FormColumns
    fFirst = 0
    fId = fFirst
    fContractor
    fMode
    fCost
    fTerm
    fMax
End Enum


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 03.06.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curLItem As man_hour_cost
    
    With listView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fId + 1).text = "Код"
            .Item(FormColumns.fContractor + 1).text = "Подрядчик"
            .Item(FormColumns.fMode + 1).text = "Режим"
            .Item(FormColumns.fCost + 1).text = "Стоимость"
            .Item(FormColumns.fTerm + 1).text = "Период"
        End With
        
        ' заполнение данными
        .ListItems.Clear
        For i = 1 To m_this.count
            Set curLItem = m_this(i)
            Set listX = .ListItems.add(, , curLItem.Id)
            For j = 1 To FormColumns.fMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormColumns.fContractor).text = curLItem.Contractor.Name
            listX.ListSubItems(FormColumns.fMode).text = curLItem.Mode.Name
            listX.ListSubItems(FormColumns.fCost).text = curLItem.Cost
            listX.ListSubItems(FormColumns.fTerm).text = curLItem.Term.StringValue
        Next i
    End With
    
    ' ширины столбцов
'    Call AppNewAutosizeColumns(listView)
    listView.ColumnHeaders(FormColumns.fContractor + 1).Width = listView.Width * 0.5
    listView.ColumnHeaders(FormColumns.fId + 1).Width = 0
    
    Set curLItem = Nothing
End Sub

Private Property Get basicListForm_Title() As String
    basicListForm_Title = "Стоимость человекочаса"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' 02.06.2022
' ----------------------------------------------------------------------------
    Set m_this = New Collection
End Sub


Public Sub initialByTerm(InTermId As Long)
' ----------------------------------------------------------------------------
' стоимости указанного месяца
' 03.06.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As man_hour_cost
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    Set m_this = New Collection
    sqlString = "get_man_hour_cost_rates_by_term"
    sqlParams.add "InTermId", InTermId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New man_hour_cost
        tempItem.add _
                    ItemId:=rst!Id, _
                    contractorId:=rst!contractor_id, _
                    termId:=rst!term_id, _
                    modeId:=rst!mode_id, _
                    CostValue:=rst!cost_sum
        
        m_this.add tempItem
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
    Set sqlParams = Nothing
End Sub


Public Sub initialCurrent()
' ----------------------------------------------------------------------------
' стоимости текущего месяца
' 02.06.2022
' ----------------------------------------------------------------------------
    Call initialByTerm(InTermId:=Terms.LastTerm.Id)
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' 02.06.2022
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = m_this.count To 1 Step -1
        m_this.remove i
    Next i
    Set m_this = Nothing
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' 02.06.2022
' ----------------------------------------------------------------------------
    count = m_this.count
End Property


Property Get Item(itemidx As Variant) As man_hour_cost
' ----------------------------------------------------------------------------
' элемент коллекции
' 02.06.2022
' ----------------------------------------------------------------------------
    Set Item = m_this.Item(itemidx)
End Property
