VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "flat_accrueds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' история начислений по квартире

Private m_FlatNo As String
Private m_FlatId As Long


Private Enum FormColumns
' ----------------------------------------------------------------------------
' столбцы формы
' 21.06.2022
' ----------------------------------------------------------------------------
    fFirst = 0
    fTerm = fFirst
    fOccId
    fFSourceName
    fServiceName
    fInSaldo
    fAccrued
    fAddeds
    fCompens
    fPeopleAccrued
    fPaid
    fOutSaldo
    fDolg
    fMax
End Enum
    

Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 21.06.2022
' ----------------------------------------------------------------------------
    Call fillListform(listView)
End Sub

Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' реализация базового класса - заголовок формы
' 21.06.2022
' ----------------------------------------------------------------------------
    basicListForm_Title = "История начислений по помещению № " & m_FlatNo
End Property


Public Sub initial(InFlatNo As String, InFlatId As Long)
' ----------------------------------------------------------------------------
' сохранение номера и кода помещения
' 21.06.2022
' ----------------------------------------------------------------------------
    m_FlatNo = InFlatNo
    m_FlatId = InFlatId
End Sub


Private Sub fillListform(listView As Object)
' ----------------------------------------------------------------------------
' заполнение списка
' 21.06.2022
' ----------------------------------------------------------------------------
    If m_FlatId = 0 Or m_FlatId = NOTVALUE Then Exit Sub
    
    Dim i As Integer
    Dim listX As listItem
    
    With listView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fAccrued + 1).text = "Начислено"
            .Item(FormColumns.fAddeds + 1).text = "Разовые"
            .Item(FormColumns.fCompens + 1).text = "Субсидия"
            .Item(FormColumns.fFSourceName + 1).text = "Источник"
            .Item(FormColumns.fInSaldo + 1).text = "Вх.сальдо"
            .Item(FormColumns.fOccId + 1).text = "Лицевой"
            .Item(FormColumns.fOutSaldo + 1).text = "Исх.сальдо"
            .Item(FormColumns.fPaid + 1).text = "Оплата"
            .Item(FormColumns.fPeopleAccrued + 1).text = "Начислено населению"
            .Item(FormColumns.fServiceName + 1).text = "Услуга"
            .Item(FormColumns.fTerm + 1).text = "Период"
            .Item(FormColumns.fDolg + 1).text = "Задолженность"
        End With
        
        Dim sqlString As String
        Dim sqlParams As New Dictionary
        Dim rst As ADODB.Recordset
        
        sqlString = "get_accrued_history_by_flat"
        sqlParams.add "InFlatId", m_FlatId
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.EOF Or rst.BOF Then GoTo cleanHandler
        ' заполнение данными
        .ListItems.Clear
        Do While Not rst.EOF
            Set listX = .ListItems.add(, , Terms(CStr(rst!OutTermId)).StringValue)
            For i = 1 To FormColumns.fMax - 1
                listX.ListSubItems.add
            Next i
            listX.ListSubItems(FormColumns.fAccrued).text = rst!outAccrued
            listX.ListSubItems(FormColumns.fAddeds).text = rst!outAdded
            listX.ListSubItems(FormColumns.fCompens).text = rst!outCompens
            If DBgetString(rst!OutInSaldo) <> "" Then
                    listX.ListSubItems(FormColumns.fDolg).text = _
                            WorksheetFunction.Max(rst!OutInSaldo - rst!OutPaid, 0)
            End If
            listX.ListSubItems(FormColumns.fFSourceName).text = rst!outFSourceName
            listX.ListSubItems(FormColumns.fInSaldo).text = DBgetString(rst!OutInSaldo)
            listX.ListSubItems(FormColumns.fOccId).text = rst!outOccId
            listX.ListSubItems(FormColumns.fOutSaldo).text = DBgetString(rst!outOutSaldo)
            listX.ListSubItems(FormColumns.fPaid).text = rst!OutPaid
            listX.ListSubItems(FormColumns.fPeopleAccrued).text = rst!outAccrued _
                    + rst!outAdded - rst!outCompens
            listX.ListSubItems(FormColumns.fServiceName).text = rst!outServiceName
            rst.MoveNext
        Loop
    End With
    
    ' ширины столбцов
    Call AppNewAutosizeColumns(listView)
    listView.ColumnHeaders(FormColumns.fAccrued + 1).Width = 0
    listView.ColumnHeaders(FormColumns.fCompens + 1).Width = 0
    listView.ColumnHeaders(FormColumns.fAddeds + 1).Width = 0
    Call AutoResizeListViewColumnHeader(listView, FormColumns.fInSaldo + 1)
    Call AutoResizeListViewColumnHeader(listView, FormColumns.fDolg + 1)
    Call AutoResizeListViewColumnHeader(listView, FormColumns.fOutSaldo + 1)
    Call AutoResizeListViewColumnHeader(listView, FormColumns.fPaid + 1)
    Call AutoResizeListViewColumnHeader(listView, FormColumns.fPeopleAccrued + 1)
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then
            rst.Close
        End If
    End If
End Sub
