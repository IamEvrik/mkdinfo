VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "old_works"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public this As Collection


Public Sub initial(BldnId As Long)
    Dim rst As New ADODB.Recordset
    Dim tmp As old_work
    
    Set this = New Collection
    
    rst.Open "SELECT * FROM old_works WHERE bldn_id = " & BldnId & " ORDER BY work_year DESC, work_name;", DBConnection.Connection
    If rst.EOF Or rst.BOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tmp = New old_work
        tmp.BldnId = BldnId
        tmp.Id = rst!Id
        tmp.wName = rst!work_name
        tmp.wNote = DBgetString(rst!Note)
        tmp.wOBF = DBgetString(rst!other_budget_flag)
        tmp.wOBN = rst!other_budget_note
        tmp.wSum = dblValue(rst!work_sum)
        tmp.wVolume = DBgetString(rst!work_volume)
        tmp.wYear = rst!work_year
        this.add tmp, CStr(tmp.Id)
        Set tmp = Nothing
        rst.MoveNext
    Loop
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
End Sub
