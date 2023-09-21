VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "subaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' информаци€ о мес€чных поступлени€х на субсчет дома
' ----------------------------------------------------------------------------

Public classBldnId As Long
Public classTermId As Long
Public classAccruedSum As Currency
Public classPaidSum As Currency
Public m_CurrentSum As Currency

