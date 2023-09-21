VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_mapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Сопоставление кодов домов наших и сторонних организаций

Private this As Collection

Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' заполнение коллекции значениями
' Last update: 01.06.2021
' ----------------------------------------------------------------------------
    Set this = New Collection
    
    Dim rst As adodb.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    sqlString = "get_bldn_mapping"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.EOF Or rst.BOF Then GoTo cleanHandler
    
    Dim tmpItem As bldn_map
    Do
        Set tmpItem = New bldn_map
        tmpItem.classOurId = rst!bldn_id
        tmpItem.classEnergoId = rst!energosbyt_bldn_id
        this.add tmpItem
        rst.MoveNext
    Loop While Not rst.EOF
    
cleanHandler:
    Set tmpItem = Nothing
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
        Set rst = Nothing
    End If
    
End Sub


Property Get GetOurIdByElectro(electroBldnId As Long) As Long
' ----------------------------------------------------------------------------
' Наш код дома по коду электричества
' Last update: 02.06.2021
' ----------------------------------------------------------------------------
    Dim colIdx As Long
    GetOurIdByElectro = NOTVALUE
    For colIdx = 1 To this.count
        If this(colIdx).classEnergoId = electroBldnId Then
            GetOurIdByElectro = this(colIdx).classOurId
            Exit Property
        End If
    Next colIdx
End Property
