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
' ��� �������
' ----------------------------------------------------------------------------

Public classId As Long          ' ���
Public classNote As String      ' ��������
Public classDate As Date        ' ���� ����
Public classValidite As Integer ' ���� �������� ����


Property Get EndDate() As Date
' ----------------------------------------------------------------------------
' ���� ��������� ���� �������
' Last update: 10.09.2020
' ----------------------------------------------------------------------------
    EndDate = DateAdd("yyyy", classValidite, classDate)
End Property


Public Sub delete()
' ----------------------------------------------------------------------------
' ��������
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "counter_certificate.delete", _
                                                            "������ �� �����"
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

