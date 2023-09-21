VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "subaccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' информация о субчетах
' ----------------------------------------------------------------------------

Implements basicListForm
Private this As Collection



Private Enum FormColumns
' -----------------------------------------------------------------------------
' поля формы поступлений на субсчет
' 29.10.2021
' -----------------------------------------------------------------------------
    fFirst = 0
    fBldnId = fFirst
    fTerm
    fAccruedSum
    fPaidSum
    fCurrentSum
    fbsMax
End Enum

Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' 02.08.2021
' ----------------------------------------------------------------------------
    Set this = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение
' 02.08.2021
' ----------------------------------------------------------------------------
    Set this = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' заголовок формы
' 02.08.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "История поступлений на субсчет"
End Property


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' заполнение формы
' 02.08.2021
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As subaccount
    
    
    With listView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fbsMax
                .add
            Next i
            .Item(FormColumns.fBldnId + 1).text = "Код"
            .Item(FormColumns.fTerm + 1).text = "Период"
            .Item(FormColumns.fPaidSum + 1).text = "Поступления"
            .Item(FormColumns.fAccruedSum + 1).text = "Начислено"
            .Item(FormColumns.fCurrentSum + 1).text = "Состояние субсчета"
        End With
        
        ' заполнение данными
        .ListItems.Clear
        For i = 1 To this.count
            Set CurLItem = this(i)
            Set listX = .ListItems.add(, , CurLItem.classBldnId)
            For j = 1 To FormColumns.fbsMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormColumns.fPaidSum).text = _
                                                        CurLItem.classPaidSum
            listX.ListSubItems(FormColumns.fAccruedSum).text = _
                                                    CurLItem.classAccruedSum
            listX.ListSubItems(FormColumns.fTerm).text = _
                                Terms(CStr(CurLItem.classTermId)).StringValue
            listX.ListSubItems(FormColumns.fCurrentSum).text = _
                                                        CurLItem.m_CurrentSum
        Next i
    End With
    
    ' ширины столбцов
    Call AppNewAutosizeColumns(listView)
    listView.ColumnHeaders(FormColumns.fBldnId + 1).Width = 0
    
    Set CurLItem = Nothing
    Set listX = Nothing

End Sub


Public Function initialByBldn(InBldnId As Long)
' ----------------------------------------------------------------------------
' получение информации по дому
' 02.08.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary, sqlString As String
    Dim tempItem As subaccount
    
    sqlString = "get_bldn_subaccount_history"
    Set sqlParams = New Dictionary
    sqlParams.add "InBldnId", InBldnId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Set this = New Collection
    Do While Not rst.EOF
        Set tempItem = New subaccount
        tempItem.classBldnId = rst!OutBldnId
        tempItem.classTermId = rst!OutTermId
        tempItem.classAccruedSum = dblValue(rst!OutAccruedSum)
        tempItem.classPaidSum = dblValue(rst!OutPaidSum)
        tempItem.m_CurrentSum = dblValue(rst!OutCurrentSum)
        
        this.add tempItem
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
End Function

