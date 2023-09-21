VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_worklist_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public wtName As String
Public wkName As String
Public cName As String
Public wDate As String
Public wSum As Currency
Public wVolume As String
Public wNote As String
Public wBldn As Long
Public wDogovor As String
Public wGWT As Long
Public wSI As String
Public wPrintFlag As Boolean
Public wFSource As String

Property Get fullVolume() As String
    fullVolume = wVolume & " " & wSI
End Property

Property Get fullWorkName() As String
    fullWorkName = wkName & IIf(wNote <> "", " (" & wNote & ")", "")
End Property
