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
' ���������� � ��������� �������
' ----------------------------------------------------------------------------

Private classId As Long                 ' ��� ����
Private classInventoryArea As Double    ' ������� ���������� ������� �� _
                                        ' ������ ����������� ��������������
Private classUseArea As Double          ' ������� ���������� ������� �� _
                                        ' ������������ �������������
Private classSurveyArea As Double       ' ������� ���������� ������� �� _
                                        ' ������ ���������
Private classBuiltUp As Double          ' ������� ���������
Private classUndeveloped As Double      ' ������������� �������
Private classHardCoatings As Double     ' ������� �������� �����
Private classDriveWays As Double        ' ������� (������ � ������� ��������)
Private classSideWalks As Double        ' �������� (������ � ������� ��������)
Private classOthers As Double           ' ������ (������ � ������� ��������)
Private classCadastralNo As String      ' ����������� ����� ���������� �������
Private classSAF As Boolean             ' ������� ����� ������������� ����
Private classFences As Boolean          ' ������� ����������
Private classBenches As Long            ' ���������� �������

Private Const SquareErrorString = _
        "����� �������� ������ �������� ������ ����� ������� ������� ��������"


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' Last update: 06.04.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' ������������� ������ �� ���� ����
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
' ��� ����
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    BldnId = classId
End Property


Property Get InventoryArea() As Double
' ----------------------------------------------------------------------------
' ������� ������� �� ������ ����������� �������������
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    InventoryArea = classInventoryArea
End Property


Property Get UseArea() As Double
' ----------------------------------------------------------------------------
' ������� ������� �� ������������ �������������
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    UseArea = classUseArea
End Property


Property Get SurveyArea() As Double
' ----------------------------------------------------------------------------
' ������� ������� �� ������ ���������
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    SurveyArea = classSurveyArea
End Property


Property Get BuiltUp() As Double
' ----------------------------------------------------------------------------
' ������� ���������
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    BuiltUp = classBuiltUp
End Property


Property Get Undeveloped() As Double
' ----------------------------------------------------------------------------
' ������������� �������
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    Undeveloped = classUndeveloped
End Property


Property Get HardCoatings() As Double
' ----------------------------------------------------------------------------
' ������ �������� �����
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    HardCoatings = classHardCoatings
End Property


Property Get DriveWays() As Double
' ----------------------------------------------------------------------------
' ������� (������ � ������ ��������)
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    DriveWays = classDriveWays
End Property


Property Get SideWalks() As Double
' ----------------------------------------------------------------------------
' �������� (������ � ������ ��������)
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    SideWalks = classSideWalks
End Property


Property Get Others() As Double
' ----------------------------------------------------------------------------
' ������ ������ ��������
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    Others = classOthers
End Property


Property Get CadastralNo() As String
' ----------------------------------------------------------------------------
' ����������� �����
' Last update: 27.04.2016
' ----------------------------------------------------------------------------
    CadastralNo = classCadastralNo
End Property


Property Get SAF() As Boolean
' ----------------------------------------------------------------------------
' ������� ����� ������������� ����
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    SAF = classSAF
End Property


Property Get Fences() As Boolean
' ----------------------------------------------------------------------------
' ������� ����������
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Fences = classFences
End Property


Property Get Benches() As Long
' ----------------------------------------------------------------------------
' ���������� �������
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Benches = classBenches
End Property


Property Let InventoryArea(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ������� �� ������ ����������� �������������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newInventory:=newValue)
End Property


Property Let UseArea(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ������� �� ������������ �������������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newUse:=newValue)
End Property


Property Let SurveyArea(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ������� �� ������ ���������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newSurvey:=newValue)
End Property


Property Let BuiltUp(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ���������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newBuilt:=newValue)
End Property


Property Let Undeveloped(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������������� �������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newUndeveloped:=newValue)
End Property


Property Let HardCoatings(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ������ ��������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newHard:=newValue)
End Property


Property Let DriveWays(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� �������� (������ � ������ ��������)
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newDrive:=newValue)
End Property


Property Let SideWalks(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ��������� (������ � ������ ��������)
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newSide:=newValue)
End Property


Property Let Others(newValue As Double)
' ----------------------------------------------------------------------------
' ��������� ������� ������ ������ ��������
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Call update(newOther:=newValue)
End Property


Property Let CadastralNo(newValue As String)
' ----------------------------------------------------------------------------
' ��������� ������������ ������
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
' ��������� ����������
' Last update: 15.05.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' ���� ��������� �� ������, �� �� ������������� �������� ������
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
' ��������� ��������
' Last update: 08.04.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
