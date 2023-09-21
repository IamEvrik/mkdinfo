VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldnLandInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' информация о земельном участке
' ----------------------------------------------------------------------------

Private classId As Long                 ' код дома
Private classInventoryArea As Double    ' площадь земельного участка по _
                                        ' данным технической инвентаризации
Private classUseArea As Double          ' площадь земельного участка по _
                                        ' фактическому использованию
Private classSurveyArea As Double       ' площадь земельного участка по _
                                        ' данным межевания
Private classBuiltUp As Double          ' площадь застройки
Private classUndeveloped As Double      ' незастроенная площадь
Private classHardCoatings As Double     ' твердые покрытия всего
Private classDriveWays As Double        ' проезды (входят в твердые покрытия)
Private classSideWalks As Double        ' тротуары (входят в твердые покрытия)
Private classOthers As Double           ' прочие (входят в твердые покрытия)
Private classCadastralNo As String      ' кадастровый номер земельного участка
Private classSAF As Boolean             ' наличие малых архитектурных форм
Private classFences As Boolean          ' наличие ограждений
Private classBenches As Long            ' количество скамеек

Private Const SquareErrorString = _
        "сумма площадей твёрдых покрытий больше общей площади твердых покрытий"


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 06.04.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' инициализация класса по коду дома
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    If itemId = NOTVALUE Then Exit Sub

    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getBuildingLandInfo"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("itemid", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    classInventoryArea = dblValue(rst!inventory_area)
    classUseArea = dblValue(rst!use_area)
    classSurveyArea = dblValue(rst!survey_area)
    classBuiltUp = dblValue(rst!builtup_area)
    classUndeveloped = dblValue(rst!undeveloped_area)
    classHardCoatings = dblValue(rst!hard_coatings)
    classDriveWays = dblValue(rst!drive_ways_hard)
    classSideWalks = dblValue(rst!side_walks_hard)
    classOthers = dblValue(rst!others_hard)
    classCadastralNo = DBgetString(rst!cadastral_no)
    classSAF = boolValue(rst!SAF)
    classFences = boolValue(rst!Fences)
    classBenches = longValue(rst!Benches)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "building_landinfo.initial", Err.Description
    End If
End Sub


Property Get BldnId() As Long
' ----------------------------------------------------------------------------
' код дома
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    BldnId = classId
End Property


Property Get InventoryArea() As Double
' ----------------------------------------------------------------------------
' площадь участка по данным технической инветаризации
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    InventoryArea = classInventoryArea
End Property


Property Get UseArea() As Double
' ----------------------------------------------------------------------------
' площадь участка по фактическому использованию
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    UseArea = classUseArea
End Property


Property Get SurveyArea() As Double
' ----------------------------------------------------------------------------
' площадь участка по данным межевания
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    SurveyArea = classSurveyArea
End Property


Property Get BuiltUp() As Double
' ----------------------------------------------------------------------------
' площадь застройки
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    BuiltUp = classBuiltUp
End Property


Property Get Undeveloped() As Double
' ----------------------------------------------------------------------------
' незастроенная площадь
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    Undeveloped = classUndeveloped
End Property


Property Get HardCoatings() As Double
' ----------------------------------------------------------------------------
' твёрдые покрытия всего
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    HardCoatings = classHardCoatings
End Property


Property Get DriveWays() As Double
' ----------------------------------------------------------------------------
' проезды (входят в твёрдые покрытия)
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    DriveWays = classDriveWays
End Property


Property Get SideWalks() As Double
' ----------------------------------------------------------------------------
' тротуары (входят в твёрдые покрытия)
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    SideWalks = classSideWalks
End Property


Property Get Others() As Double
' ----------------------------------------------------------------------------
' прочие твёрдые покрытия
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    Others = classOthers
End Property


Property Get CadastralNo() As String
' ----------------------------------------------------------------------------
' кадастровый номер
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    CadastralNo = classCadastralNo
End Property


Property Get SAF() As Boolean
' ----------------------------------------------------------------------------
' наличие малых архитектурных форм
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    SAF = classSAF
End Property


Property Get Fences() As Boolean
' ----------------------------------------------------------------------------
' наличие ограждений
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Fences = classFences
End Property


Property Get Benches() As Long
' ----------------------------------------------------------------------------
' количество скамеек
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Benches = classBenches
End Property


Property Let InventoryArea(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади участка по данным технической инветаризации
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newInventory:=newValue)
End Property


Property Let UseArea(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади участка по фактическому использованию
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newUse:=newValue)
End Property


Property Let SurveyArea(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади участка по данным межевания
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newSurvey:=newValue)
End Property


Property Let BuiltUp(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади застройки
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newBuilt:=newValue)
End Property


Property Let Undeveloped(newValue As Double)
' ----------------------------------------------------------------------------
' изменение незастроенной площади
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newUndeveloped:=newValue)
End Property


Property Let HardCoatings(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади твёрдых покрытий
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newHard:=newValue)
End Property


Property Let DriveWays(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади проездов (входят в твёрдые покрытия)
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newDrive:=newValue)
End Property


Property Let SideWalks(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади тротуаров (входят в твёрдые покрытия)
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newSide:=newValue)
End Property


Property Let Others(newValue As Double)
' ----------------------------------------------------------------------------
' изменение площади прочих твёрдых покрытий
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newOther:=newValue)
End Property


Property Let CadastralNo(newValue As String)
' ----------------------------------------------------------------------------
' изменение кадастрового номера
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newCadastral:=newValue)
End Property


Public Sub update(Optional newInventory As Double = NOTVALUE, _
                    Optional newUse As Double = NOTVALUE, _
                    Optional newSurvey As Double = NOTVALUE, _
                    Optional newBuilt As Double = NOTVALUE, _
                    Optional newUndeveloped As Double = NOTVALUE, _
                    Optional newHard As Double = NOTVALUE, _
                    Optional newDrive As Double = NOTVALUE, _
                    Optional newSide As Double = NOTVALUE, _
                    Optional newOther As Double = NOTVALUE, _
                    Optional newCadastral As String = NOTSTRING, _
                    Optional newFences As Variant, _
                    Optional newSAF As Variant, _
                    Optional newBenches As Long = NOTVALUE)
' ----------------------------------------------------------------------------
' изменение информации
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If newInventory = NOTVALUE Then newInventory = classInventoryArea
    If newUse = NOTVALUE Then newUse = classUseArea
    If newSurvey = NOTVALUE Then newSurvey = classSurveyArea
    If newBuilt = NOTVALUE Then newBuilt = classBuiltUp
    If newUndeveloped = NOTVALUE Then newUndeveloped = classUndeveloped
    If newDrive = NOTVALUE Then newDrive = classDriveWays
    If newHard = NOTVALUE Then newHard = classHardCoatings
    If newSide = NOTVALUE Then newSide = classSideWalks
    If newOther = NOTVALUE Then newOther = classOthers
    If StrComp(newCadastral, NOTSTRING, vbBinaryCompare) = 0 Then _
                                            newCadastral = classCadastralNo
    If IsMissing(newSAF) Then newSAF = classSAF
    If IsMissing(newFences) Then newFences = classFences
    If newBenches = NOTVALUE Then newBenches = classBenches
                                            
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "change_bldn_land_info"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = classId
    cmd.Parameters("newinv").Value = newInventory
    cmd.Parameters("newuse").Value = newUse
    cmd.Parameters("newsurv").Value = newSurvey
    cmd.Parameters("newbuilt").Value = newBuilt
    cmd.Parameters("newundev").Value = newUndeveloped
    cmd.Parameters("newhard").Value = newHard
    cmd.Parameters("newdrive").Value = newDrive
    cmd.Parameters("newside").Value = newSide
    cmd.Parameters("newother").Value = newOther
    cmd.Parameters("newcadastr").Value = newCadastral
    cmd.Parameters("newsaf").Value = newSAF
    cmd.Parameters("newfences").Value = newFences
    cmd.Parameters("newbenches").Value = newBenches
    Set rst = cmd.Execute
        
    classInventoryArea = newInventory
    classUseArea = newUse
    classSurveyArea = newSurvey
    classBuiltUp = newBuilt
    classUndeveloped = newUndeveloped
    classHardCoatings = newHard
    classSideWalks = newSide
    classOthers = newOther
    classCadastralNo = newCadastral
    classDriveWays = newDrive
    classSAF = newSAF
    classFences = newFences
    classBenches = newBenches
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldnlandinfo_class.update_dogovor", Err.Description
    End If
End Sub

 Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
