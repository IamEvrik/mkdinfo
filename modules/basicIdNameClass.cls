VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "basicIdNameClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Name() As String
End Property

Public Property Get Id() As Long
End Property

Public Sub initial(ItemId As Long)
End Sub

Public Sub delete()
End Sub

Public Sub showForm(isChange As Boolean)
End Sub
