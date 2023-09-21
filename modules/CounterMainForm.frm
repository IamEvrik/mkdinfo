VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CounterMainForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12900
   OleObjectBlob   =   "CounterMainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CounterMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Debug.Print Node.text

End Sub

Private Sub UserForm_Activate()
    With Me.TreeView1
        Dim tmp As New address_md_list
        Dim tmp1 As New address_village_list
        Dim tmp2 As New address_street_list
        Dim i As Long, j As Long, k As Long
        
        For i = 1 To tmp.count
            .Nodes.add , , "mdid: " & tmp(i).Id, tmp(i).Name
            For j = 1 To tmp1.count
                If tmp1(j).Municipal_district.Id = tmp(i).Id Then
                    .Nodes.add "mdid: " & tmp(i).Id, tvwChild, "vlid: " & tmp1(j).Id, tmp1(j).Name
                    For k = 1 To tmp2.count
                        If tmp2(k).Village.Id = tmp1(j).Id Then
                            .Nodes.add "vlid: " & tmp1(j).Id, tvwChild, tmp1(j).Name & tmp2(k).Name, tmp2(k).Name
                        End If
                    Next k
                    .Nodes("vlid: " & tmp1(j).Id).Expanded = False
                    
                End If
            Next j
            .Nodes("mdid: " & tmp(i).Id).Expanded = False
        Next i
        
    End With
End Sub

