VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "old_work"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Id As Long
Public BldnId As Long
Public wName As String
Public wYear As Long
Public wVolume As String
Public wSum As Currency
Public wNote As String
Public wOBF As Boolean
Public wOBN As String

Public Sub delete()
    If Id <> 0 And Id <> NOTVALUE Then
        Dim rst As New ADODB.Recordset
        rst.Open "DELETE FROM old_works WHERE ID=" & Id & ";", DBConnection.Connection, adOpenDynamic, adLockOptimistic
        If Not rst Is Nothing Then
            If rst.State = adStateOpen Then rst.Close
        End If
        Set rst = Nothing
    End If
End Sub
