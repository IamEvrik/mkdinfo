VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "works_materials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ��������� �����
' ----------------------------------------------------------------------------

Public Id As Long
Public MaintenanceWorkId As Long
Public MaterialId As Long
Public MaterialNote As String
Public MaterialCost As Currency
Public MaterialCount As Double
Public MaterialSi As String


Property Get MaterialCountJSON() As String
' ----------------------------------------------------------------------------
' ���������� ��� ������ � json (����������� �����)
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    MaterialCountJSON = Replace(CStr(MaterialCount), ",", ".")
End Property
    
    
Property Get MaterialCostJSON() As String
' ----------------------------------------------------------------------------
' ���� ��� ������ � json (����������� �����)
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    MaterialCostJSON = Replace(CStr(MaterialCost), ",", ".")
End Property


Property Get MaterialSum() As Currency
' ----------------------------------------------------------------------------
' ��������� ����������
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    MaterialSum = MaterialCount * MaterialCost
End Property


Property Get MaterialNoteJSON() As String
' ----------------------------------------------------------------------------
' ���������� � ��������� ��� JSON
' Last update: 24.09.2020
' ----------------------------------------------------------------------------
    MaterialNoteJSON = EscapeJSON(MaterialNote)
End Property
