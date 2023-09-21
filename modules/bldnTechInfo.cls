VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldnTechInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' техническая информация МКД
' ----------------------------------------------------------------------------

Private classId As Long                     ' код дома
Private classFloorMin As Integer            ' этажей минимальное
Private classFloorMax As Integer            ' этажей максимальное
Private classVaults As Integer              ' подвалов
Private classEntrance As Integer            ' подъездов
Private classBuiltYear As Integer           ' год постройки
Private classCommissioningYear As Integer   ' год ввода в эксплуатацию
Private classDepreciation As Double         ' износ
Private classAtticSquare As Double          ' площадь чердаков
Private classVaultSquare As Double          ' площадь подвалов
Private classStairsSquare As Double         ' площадь лестничных маршев и
                                            ' площадок
Private classCorridorSquare As Double       ' площадь коридоров МОП
Private classOtherSquare As Double          ' площадь иных помещений МОП
Private classWallMaterial As Long           ' материал стен
Private classStairs As Integer              ' количество лестниц
Private classStructuralVolume As Double     ' строительный объем
Private classHasDoorPhone As Boolean        ' наличие домофона
Private classDoorPhoneComment As String     ' комментарий к домофону
' Наличие ОДПУ
Private classHasHotWater As Boolean         ' ГВС
Private classHasHeating As Boolean          ' Отопление
Private classHasColdWater As Boolean        ' ХВС
Private classHasCommon As Boolean           ' теплоснабжение общий
Private classHasElectro As Boolean          ' электроэнергия
Private classHasThermoregulator As Boolean  ' погодозависимая автоматика
Private classHasDoorCloser As Boolean       ' доводчик

Private c_ProcessedSurface As processed_surface ' обрабатываемые поверхности


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
' ---------------------------------------------------------------------------
' инициализация класса по коду
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    If ItemId = NOTVALUE Then Exit Sub

    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_bldn_tech_info"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemid").Value = ItemId
    
    Set rst = cmd.Execute
    
    classId = ItemId
    classFloorMin = longValue(rst!floor_min)
    classFloorMax = longValue(rst!floor_max)
    classVaults = longValue(rst!vaults)
    classEntrance = longValue(rst!entrances)
    classBuiltYear = longValue(rst!built_year)
    classCommissioningYear = longValue(rst!commissioning_year)
    classDepreciation = dblValue(rst!Depreciation)
    classAtticSquare = dblValue(rst!attic_square)
    classVaultSquare = dblValue(rst!vaults_square)
    classStairsSquare = dblValue(rst!stairs_square)
    classCorridorSquare = dblValue(rst!corridor_square)
    classOtherSquare = dblValue(rst!other_square)
    classWallMaterial = longValue(rst!wallmater_id)
    classStairs = longValue(rst!stairs)
    classStructuralVolume = dblValue(rst!structural_volume)
    classHasHotWater = boolValue(rst!has_odpu_hotwater)
    classHasHeating = boolValue(rst!has_odpu_heating)
    classHasColdWater = boolValue(rst!has_odpu_coldwater)
    classHasCommon = boolValue(rst!has_odpu_common)
    classHasElectro = boolValue(rst!has_odpu_electro)
    classHasDoorPhone = boolValue(rst!has_doorphone)
    classDoorPhoneComment = DBgetString(rst!doorphone_comment)
    classHasThermoregulator = boolValue(rst!has_thermoregulator)
    classHasDoorCloser = boolValue(rst!has_doorcloser)
    
    Set c_ProcessedSurface = New processed_surface
    With c_ProcessedSurface
        .SquareBanisters = dblValue(rst!square_banisters)
        .SquareDoorHandles = dblValue(rst!square_doorhandles)
        .SquareDoors = dblValue(rst!square_doors)
        .SquareMailBoxes = dblValue(rst!square_mailboxes)
        .SquareWindowSills = dblValue(rst!square_windowsills)
        .SquareRadiators = dblValue(rst!square_radiators)
    End With
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "building_techifo.initial", Err.Description
    End If
End Sub


Property Get BldnId() As Long
' ----------------------------------------------------------------------------
' код дома
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    BldnId = classId
End Property


Property Get FloorMin() As Integer
' ----------------------------------------------------------------------------
' количество этажей минимальное
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    FloorMin = classFloorMin
End Property


Property Get FloorMax() As Integer
' ----------------------------------------------------------------------------
' количество этажей максимальное
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    FloorMax = classFloorMax
End Property


Property Get VaultsCount() As Integer
' ----------------------------------------------------------------------------
' количество подвалов
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    VaultsCount = classVaults
End Property


Property Get EntranceCount() As Integer
' ----------------------------------------------------------------------------
' количество подъездов
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    EntranceCount = classEntrance
End Property


Property Get StairsCount() As Integer
' ----------------------------------------------------------------------------
' количество лестниц
' Last update: 15.05.2016
' ----------------------------------------------------------------------------
    StairsCount = classStairs
End Property

Property Get BuiltYear() As Integer
' ----------------------------------------------------------------------------
' год постройки
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    BuiltYear = classBuiltYear
End Property


Property Get CommissioningYear() As Integer
' ----------------------------------------------------------------------------
' год ввода в эксплутацию
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    CommissioningYear = classCommissioningYear
End Property


Property Get Depreciation() As Double
' ----------------------------------------------------------------------------
' износ
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    Depreciation = classDepreciation
End Property


Property Get AtticSquare() As Double
' ----------------------------------------------------------------------------
' площадь чердаков
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    AtticSquare = classAtticSquare
End Property


Property Get VaultSquare() As Double
' ----------------------------------------------------------------------------
' площадь подвалов
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    VaultSquare = classVaultSquare
End Property


Property Get StairsSquare() As Double
' ----------------------------------------------------------------------------
' площадь лестничных маршей и площадок
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    StairsSquare = classStairsSquare
End Property


Property Get CorridorSquare() As Double
' ----------------------------------------------------------------------------
' площадь корридоров МОП
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    CorridorSquare = classCorridorSquare
End Property


Property Get OtherSquare() As Double
' ----------------------------------------------------------------------------
' площадь иных помещений МОП
' Last update: 17.11.2016
' ----------------------------------------------------------------------------
    OtherSquare = classOtherSquare
End Property


Property Get StructuralVolume() As Double
' ----------------------------------------------------------------------------
' строительный объем
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    StructuralVolume = classStructuralVolume
End Property


Property Get WallMaterial() As wallmaterial_class
' ----------------------------------------------------------------------------
' материал стен
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    Set WallMaterial = wallmaterial_list(CStr(classWallMaterial))
End Property


Property Get MOPSquare() As Double
' ----------------------------------------------------------------------------
' площадь МОП
' Last update: 03.05.2016
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        MOPSquare = NOTVALUE
    Else
        MOPSquare = classCorridorSquare + classStairsSquare + classOtherSquare
    End If
End Property


Property Get HasDoorPhone() As Boolean
' ----------------------------------------------------------------------------
' наличие домофона
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasDoorPhone = NOTVALUE
    Else
        HasDoorPhone = classHasDoorPhone
    End If
End Property


Property Get DoorPhoneComment() As String
' ----------------------------------------------------------------------------
' комментарий к домофону
' Last update: 27.01.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        DoorPhoneComment = NOTSTRING
    Else
        DoorPhoneComment = classDoorPhoneComment
    End If
End Property


Property Get HasDoorCloser() As Boolean
' ----------------------------------------------------------------------------
' наличие доводчика
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasDoorCloser = NOTVALUE
    Else
        HasDoorCloser = classHasDoorCloser
    End If
End Property


Property Get HasOdpuHotWater() As Boolean
' ----------------------------------------------------------------------------
' ОДПУ горячей воды
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasOdpuHotWater = NOTVALUE
    Else
        HasOdpuHotWater = classHasHotWater
    End If
End Property


Property Get HasOdpuColdWater() As Boolean
' ----------------------------------------------------------------------------
' ОДПУ ХВС
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasOdpuColdWater = NOTVALUE
    Else
        HasOdpuColdWater = classHasColdWater
    End If
End Property


Property Get HasOdpuHeating() As Boolean
' ----------------------------------------------------------------------------
' ОДПУ отопления
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasOdpuHeating = NOTVALUE
    Else
        HasOdpuHeating = classHasHeating
    End If
End Property


Property Get HasOdpuCommon() As Boolean
' ----------------------------------------------------------------------------
' ОДПУ общий теплоснабжение
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasOdpuCommon = NOTVALUE
    Else
        HasOdpuCommon = classHasCommon
    End If
End Property


Property Get HasOdpuElectro() As Boolean
' ----------------------------------------------------------------------------
' ОДПУ электроэнергия
' Last update: 10.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasOdpuElectro = NOTVALUE
    Else
        HasOdpuElectro = classHasElectro
    End If
End Property


Property Get HasThermoregulator() As Boolean
' ----------------------------------------------------------------------------
' наличие погодозависимой автоматики
' Last update: 26.02.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        HasThermoregulator = NOTVALUE
    Else
        HasThermoregulator = classHasThermoregulator
    End If
End Property


Property Get ProcessedSurface() As processed_surface
' ----------------------------------------------------------------------------
' обрабатываемые поверхности
' Last update: 23.04.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        Set ProcessedSurface = c_ProcessedSurface
    End If
End Property


Public Sub update(newFloorMin As Long, newFloorMax As Long, _
                    newVaultsCount As Long, newEntranceCount As Long, _
                    newStairsCount As Long, newBuiltYear As Long, _
                    newCommissioningYear As Long, newDepreciation As Double, _
                    newAtticSquare As Double, newVaultSquare As Double, _
                    newStairsSquare As Double, newCorridorSquare As Double, _
                    newOtherSquare As Double, newStructuralVolume As Double, _
                    newWallmaterial As Long, newHasHotWater As Variant, _
                    newHasColdWater As Variant, newHasCommon As Variant, _
                    newHasHeating As Variant, newHasElectro As Variant, _
                    newHasDoorPhone As Variant, newDoorPhoneComment As String, _
                    newHasThermoregulator As Boolean, _
                    newSquareBanisters As Double, newSquareDoors As Double, _
                    newSquareWindowSills As Double, newSquareDoorHandles As Double, _
                    newSquareMailBoxes As Double, newSquareRadiators As Double, _
                    newHasDoorCloser As Boolean _
                    )
' ----------------------------------------------------------------------------
' изменение информации
' Last update: 11.11.2020
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    Dim sqlText As String
    
    On Error GoTo errHandler
    
    sqlText = "change_bldn_tech_info"
    sqlParams.add "itemId", classId
    sqlParams.add "newfmin", newFloorMin
    sqlParams.add "newfmax", newFloorMax
    sqlParams.add "newvaults", newVaultsCount
    sqlParams.add "newentrance", newEntranceCount
    sqlParams.add "newstairs", newStairsCount
    sqlParams.add "newbuilt", newBuiltYear
    sqlParams.add "newcommissioning", newCommissioningYear
    sqlParams.add "newdepreciation", newDepreciation
    sqlParams.add "newatticsq", newAtticSquare
    sqlParams.add "newvaultssq", newVaultSquare
    sqlParams.add "newstairssq", newStairsSquare
    sqlParams.add "newcorridorsq", newCorridorSquare
    sqlParams.add "newothersq", newOtherSquare
    sqlParams.add "newwall", newWallmaterial
    sqlParams.add "newstrvolume", newStructuralVolume
    sqlParams.add "newhashw", newHasHotWater
    sqlParams.add "newhascw", newHasColdWater
    sqlParams.add "newhaselectro", newHasElectro
    sqlParams.add "newhascommon", newHasCommon
    sqlParams.add "newhasheating", newHasHeating
    sqlParams.add "newhasdoorphone", newHasDoorPhone
    sqlParams.add "newdoorphonecomment", newDoorPhoneComment
    sqlParams.add "newhasthermoregulator", newHasThermoregulator
    sqlParams.add "InBanisterSq", newSquareBanisters
    sqlParams.add "InDoorsSq", newSquareDoors
    sqlParams.add "InWindowSillsSq", newSquareWindowSills
    sqlParams.add "InDoorHandlesSq", newSquareDoorHandles
    sqlParams.add "InMailBoxesSq", newSquareMailBoxes
    sqlParams.add "InRadiatorsSq", newSquareRadiators
    sqlParams.add "InHasDoorCloser", newHasDoorCloser
    
    Call DBConnection.RunQuery(sqlText, sqlParams)
    
    classFloorMin = newFloorMin
    classFloorMax = newFloorMax
    classVaults = newVaultsCount
    classEntrance = newEntranceCount
    classStairs = newStairsCount
    classBuiltYear = newBuiltYear
    classCommissioningYear = newCommissioningYear
    classDepreciation = newDepreciation
    classAtticSquare = newAtticSquare
    classVaultSquare = newVaultSquare
    classStairsSquare = newStairsSquare
    classCorridorSquare = newCorridorSquare
    classOtherSquare = newOtherSquare
    classWallMaterial = newWallmaterial
    classStructuralVolume = newStructuralVolume
    classHasElectro = newHasElectro
    classHasHotWater = newHasHotWater
    classHasHeating = newHasHeating
    classHasCommon = newHasCommon
    classHasColdWater = newHasColdWater
    classHasDoorPhone = newHasDoorPhone
    classDoorPhoneComment = newDoorPhoneComment
    classHasThermoregulator = newHasThermoregulator
    classHasDoorCloser = newHasDoorCloser
    c_ProcessedSurface.SquareBanisters = newSquareBanisters
    c_ProcessedSurface.SquareDoorHandles = newSquareDoorHandles
    c_ProcessedSurface.SquareDoors = newSquareDoors
    c_ProcessedSurface.SquareMailBoxes = newSquareMailBoxes
    c_ProcessedSurface.SquareWindowSills = newSquareWindowSills
    c_ProcessedSurface.SquareRadiators = newSquareRadiators
        
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldntechinfo_class.update", Err.Description
    End If
End Sub


 Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
