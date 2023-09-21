VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_form_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Id() As Long
End Property

Public Property Get Name() As String
End Property

Public Sub initial(ItemId As Long)
End Sub

Public Sub delete()
End Sub

Public Sub showForm(isChange As Boolean)
End Sub

Property Get isFormFill(ByRef curForm As UserForm) As Boolean
End Property

Public Sub update(ByRef cuForm As UserForm, isChange As Boolean)
End Sub
