VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "counter_certificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' Акт допуска
' ----------------------------------------------------------------------------

Public classId As Long          ' код
Public classNote As String      ' описание
Public classDate As Date        ' дата акта
Public classValidite As Integer ' срок действия акта


Property Get EndDate() As Date
' ----------------------------------------------------------------------------
' Дата окончания акта допуска
' Last update: 10.09.2020
' ----------------------------------------------------------------------------
    EndDate = DateAdd("yyyy", classValidite, classDate)
End Property


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "counter_certificate.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", classId
    
    DBConnection.RunQuery "delete_counter_certificate", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "counter_certificate.delete", Err.Description
    End If
    
End Sub

