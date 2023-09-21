VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_chairmans_sign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' подписи комендантов
' ----------------------------------------------------------------------------

Private m_this As New Collection


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 16.10.2022
' ----------------------------------------------------------------------------
    Call fillListView(listView)
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' реализация базового класса - название формы
' 16.10.2022
' ----------------------------------------------------------------------------
    basicListForm_Title = "Подписи комендантов"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' конструктор по умолчанию
' 16.10.2022
' ----------------------------------------------------------------------------
End Sub


Public Sub reload(InBldnId As Long)
' ----------------------------------------------------------------------------
' обновление коллекции
' 16.10.2022
' ----------------------------------------------------------------------------
    Call initial(InBldnId)
End Sub


Public Sub initial(InBldnId As Long)
' ----------------------------------------------------------------------------
' получение списка
' 20.10.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As bldn_chairman_sign
    
    Set m_this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_chairmans_signature_in_bldn"
    sqlParams.add "InBldnId", InBldnId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New bldn_chairman_sign
        tempItem.add InBldnId:=InBldnId, _
                InTermId:=rst!begin_term, _
                InOwnerName:=DBgetString(rst!signature_owner)
        tempItem.hasSign = Not IsNull(rst!Sign)
        m_this.add tempItem, CStr(tempItem.BeginTerm.Id)
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


Public Sub fillListView(listView As Object)
' ----------------------------------------------------------------------------
' заполнение формы
' 19.10.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curLItem As bldn_chairman_sign
    
    With listView
        .Visible = False
        
        .View = lvwReport       ' в виде таблицы
        .FullRowSelect = True   ' выделяется вся строка
        .LabelEdit = lvwManual  ' запрет изменения значений в самом ListView
        .Gridlines = True       ' сетка
        
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormChairmanSignColumns.fcsMax
                .add
            Next i
            .Item(FormChairmanSignColumns.fcsBeginTerm + 1).text = "Начало действия"
            .Item(FormChairmanSignColumns.fcsBldnId + 1).text = "Код дома"
            .Item(FormChairmanSignColumns.fcsTermId + 1).text = "Код периода"
            .Item(FormChairmanSignColumns.fcsOwnerName + 1).text = "Подписант"
            .Item(FormChairmanSignColumns.fcsHasSign + 1).text = "Имеется подпись"
        End With
        
        ' заполнение данными
        .ListItems.Clear
        For i = 1 To m_this.count
            Set curLItem = m_this(i)
            Set listX = .ListItems.add(, , curLItem.BeginTerm.StringValue)
            For j = 1 To FormChairmanSignColumns.fcsMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormChairmanSignColumns.fcsOwnerName).text = curLItem.OwnerName
            listX.ListSubItems(FormChairmanSignColumns.fcsTermId).text = curLItem.BeginTerm.Id
            listX.ListSubItems(FormChairmanSignColumns.fcsBldnId).text = curLItem.BldnId
            listX.ListSubItems(FormChairmanSignColumns.fcsHasSign).text = BoolToYesNo(curLItem.hasSign)
        Next i
                
        ' ширины столбцов
        Call AppNewAutosizeColumns(listView)
        listView.ColumnHeaders(FormChairmanSignColumns.fcsBldnId + 1).Width = 0
        listView.ColumnHeaders(FormChairmanSignColumns.fcsTermId + 1).Width = 0
                
        .Visible = True
        
        ' эти две строки лечат убегание ListView из своего положения
        .Visible = False
        .Visible = True
    End With
        
    Set curLItem = Nothing

End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' 11.04.2022
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = m_this.count To 1 Step -1
        m_this.remove i
    Next i
    Set m_this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' 11.04.2022
' ----------------------------------------------------------------------------
    count = m_this.count
End Property


Property Get Item(itemidx As Variant) As bldn_common_property
' ----------------------------------------------------------------------------
' элемент коллекции
' 11.04.2022
' ----------------------------------------------------------------------------
    Set Item = m_this.Item(itemidx)
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' ----------------------------------------------------------------------------
' для итерации по списку
' 05.05.2022
' ----------------------------------------------------------------------------
    Set NewEnum = m_this.[_NewEnum]
End Property
