VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "contractor_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' список подрядчиков (статический)
' ----------------------------------------------------------------------------

Private this As New Collection



Private Enum FormColumns
' ----------------------------------------------------------------------------
' столбцы формы подрядчиков
' 17.09.2019
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fName
    fDirectorPosition
    fDirector
    fBldnContractor
    fIsUsing
    fMax
End Enum

Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' реализация базового класса - заполнение формы
' 17.09.2021
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As contractor_class
    
    
    With listView
        ' заголовки столбцов
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fBldnContractor + 1).text = "Основной"
            .Item(FormColumns.fDirector + 1).text = "Директор"
            .Item(FormColumns.fDirectorPosition + 1).text = _
                                                        "Должность директора"
            .Item(FormColumns.fID + 1).text = "Код"
            .Item(FormColumns.fName + 1).text = "Название"
            .Item(FormColumns.fIsUsing + 1).text = "Активен"
        End With
        
        ' заполнение данными
        .ListItems.Clear
        contractor_list.reload
        For i = 1 To contractor_list.count
            Set CurLItem = contractor_list(i)
            Set listX = .ListItems.add(, , CurLItem.Id)
            For j = 1 To FormColumns.fMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormColumns.fBldnContractor).text = _
                    BoolToYesNo(CurLItem.BldnContractor)
            listX.ListSubItems(FormColumns.fDirector).text = _
                    CurLItem.Director
            listX.ListSubItems(FormColumns.fDirectorPosition).text = _
                    CurLItem.DirectorPosition
            listX.ListSubItems(FormColumns.fName).text = CurLItem.Name
            listX.ListSubItems(FormColumns.fIsUsing).text = _
                    BoolToYesNo(CurLItem.isUsing)
        Next i
    End With
    
    ' ширины столбцов
    Call AppNewAutosizeColumns(listView)
    
    Set CurLItem = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' реализация базового класса - название формы
' 17.09.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "Подрядные организации"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' Last update: 10.10.2017
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление коллекции
' Last update: 15.02.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' получение списка
' 17.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As contractor_class
    
    Set this = New Collection
    
    sqlString = "get_contractors"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New contractor_class
        tempItem.add _
            Id:=rst!Id, _
            Name:=rst!Name, _
            Director:=DBgetString(rst!Director), _
            DirectorPosition:=DBgetString(rst!director_position), _
            BldnContractor:=rst!bldn_contractor, _
            setIsUsing:=rst!is_using
        
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set sqlParams = Nothing
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 30.04.2016
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
' Last update: 01.03.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As contractor_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
