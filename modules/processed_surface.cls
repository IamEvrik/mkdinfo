VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "processed_surface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' площади обрабатываемых поверхностей
' класс дл€ использовани€ в bldnTechInfo

Private c_Banister As Double        ' перила
Private c_Door As Double            ' двери
Private c_WindowSill As Double      ' подоконники
Private c_DoorHandle As Double      ' ручки
Private c_MailBox As Double         ' почтовые €щики
Private c_Radiator As Double        ' радиаторы ћќѕ

Property Get SquareBanisters() As Double
    SquareBanisters = c_Banister
End Property


Property Let SquareBanisters(ByVal newValue As Double)
    c_Banister = newValue
End Property


Property Get SquareDoors() As Double
    SquareDoors = c_Door
End Property


Property Let SquareDoors(ByVal newValue As Double)
    c_Door = newValue
End Property


Property Get SquareWindowSills() As Double
    SquareWindowSills = c_WindowSill
End Property


Property Let SquareWindowSills(ByVal newValue As Double)
    c_WindowSill = newValue
End Property


Property Get SquareDoorHandles() As Double
    SquareDoorHandles = c_DoorHandle
End Property


Property Let SquareDoorHandles(ByVal newValue As Double)
    c_DoorHandle = newValue
End Property


Property Get SquareMailBoxes() As Double
    SquareMailBoxes = c_MailBox
End Property


Property Let SquareMailBoxes(ByVal newValue As Double)
    c_MailBox = newValue
End Property


Property Get SquareRadiators() As Double
    SquareRadiators = c_Radiator
End Property


Property Let SquareRadiators(ByVal newValue As Double)
    c_Radiator = newValue
End Property

