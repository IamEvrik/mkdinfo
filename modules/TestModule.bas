Attribute VB_Name = "TestModule"
Option Explicit

Sub testClass()
    Dim tstlist As New work_maintenance
    MsgBox tstlist.PrintFlag
End Sub

Sub testForm()
    With ListForm
        .formType = lftHCounterPartTypes
        .Show
    End With
End Sub

Sub one()
    Dim test As basic_class
    Set test = New hcounter_part_type
    Call test.create("Меркурий 230 АМ-02")
    Set test = Nothing
End Sub

Sub deletemodule(df As Boolean)

    Dim curModule As VBComponent

    For Each curModule In ThisWorkbook.VBProject.VBComponents
        If curModule.Name = "SOME_NAME" Then
            ThisWorkbook.VBProject.VBComponents.remove curModule
        End If
    Next curModule
End Sub
