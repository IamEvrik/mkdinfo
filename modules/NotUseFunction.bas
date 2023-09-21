Attribute VB_Name = "NotUseFunction"
Option Explicit

Sub ImportWorksFromPlan()
' ----------------------------------------------------------------------------
' �������� ����� �� ����� (�������� �� xml)
' Last update: 06.02.2018
' ----------------------------------------------------------------------------
    Dim importName As String
    Dim i As Long
    Dim xml As Object, wNodes As Object, curXmlWork As Object
    Dim wterms As New terms
    Dim curWork As work_class
    Dim myDbHandle As Long
    Dim errFName As String, errFS As Object, errFile As Object
    Dim logstr As String
    
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    importName = Application.GetOpenFilename("xml ���� ������ (*.xml),*.xml", _
                        Title:="�������� ����")
    
    If Not xml.Load(importName) Then
        MsgBox "�� ������ ���� ��� �������"
        Exit Sub
    End If
    
    On Error GoTo errHandler
    errFName = importName & ".log"
    Set errFS = CreateObject("Scripting.FileSystemObject")
    Set errFile = errFS.opentextfile(errFName, 8, True)
    
    errFile.writeline (vbCrLf & "--------------------------------------------------")
    errFile.writeline (CStr(Now) & " ������ �������� ����� " & importName)
    
    If StrComp(xml.SelectSingleNode("/works").Attributes.getNamedItem("version").text, _
                IMPORT_FROM_PLAN_VERSION) <> 0 Then
        MsgBox "������������ ������ ����� � ��������", vbExclamation, "������"
        errFile.writeline ("������������ ������ �����: " & _
                        xml.SelectSingleNode("/works").Attributes.getNamedItem("version").text & _
                        " - " & IMPORT_FROM_PLAN_VERSION)
        Exit Sub
    End If
    
    If Not xml.SelectSingleNode("/works").Attributes.getNamedItem("status") Is Nothing Then
        MsgBox "������ ���� ��� ��� ��������", vbExclamation, "������"
        errFile.writeline ("���� ��� �������� �����")
        Exit Sub
    End If
    
    Set wNodes = xml.SelectNodes("works/work")
    ' ����������� ��� � xml-�����
    For Each curXmlWork In wNodes
        For i = wterms.count To 1 Step -1
            If curXmlWork.SelectSingleNode("work_date").text >= wterms(i).classBeginDate And _
                        curXmlWork.SelectSingleNode("work_date").text <= wterms(i).classEndDate Then
                curXmlWork.SelectSingleNode("work_date").text = wterms(i).Id
                Exit For
            End If
        Next i
    Next curXmlWork
    
    ' �������� �����
    myDbHandle = ConnectSqlite3(SQLITE_OPEN_READWRITE)
    ' ������ ����������
    SQLite3ExecuteNonQuery myDbHandle, "BEGIN TRANSACTION"
    For Each curXmlWork In wNodes
        Set curWork = New work_class
        curWork.create BldnId:=curXmlWork.SelectSingleNode("bldn_id").text, _
                        gwtId:=curXmlWork.SelectSingleNode("gwt_id").text, _
                        workKindID:=curXmlWork.SelectSingleNode("workkind_id").text, _
                        WorkDate:=curXmlWork.SelectSingleNode("work_date").text, _
                        workSum:=curXmlWork.SelectSingleNode("work_sum").text, _
                        Si:=curXmlWork.SelectSingleNode("si").text, _
                        workVolume:=curXmlWork.SelectSingleNode("volume").text, _
                        workNote:=curXmlWork.SelectSingleNode("note").text, _
                        contractorId:=curXmlWork.SelectSingleNode("contractor_id").text, _
                        mcId:=curXmlWork.SelectSingleNode("mc_id").text, _
                        Dogovor:=curXmlWork.SelectSingleNode("dogovor").text, _
                        PrintFlag:=True, _
                        dbHandle:=myDbHandle
        logstr = Join(Array(curXmlWork.SelectSingleNode("address").text, _
                            curXmlWork.SelectSingleNode("work_name").text), " ")
        errFile.writeline ("success: " & logstr)
    Next curXmlWork
    ' ������ ����������
    SQLite3ExecuteNonQuery myDbHandle, "COMMIT TRANSACTION"
    MsgBox "������ ��������"
    
    On Error GoTo 0
    xml.SelectSingleNode("/works").setattribute "status", "done"
    xml.Save importName
    
    GoTo cleanHandler
errHandler:
    ' ����� ����������
    logstr = ""
    If myDbHandle <> 0 Then
        logstr = SQLite3ErrMsg(myDbHandle)
        SQLite3ExecuteNonQuery myDbHandle, "ROLLBACK TRANSACTION"
    End If
    If errFile Is Nothing Then
        MsgBox Err.Description
    Else
        errFile.writeline logstr & " | " & Err.Description
        MsgBox "��� �������� ��������� ������"
    End If
    GoTo cleanHandler
    
cleanHandler:
    If myDbHandle <> 0 Then Call ConnectSqlite3Close(myDbHandle)
    If Not errFile Is Nothing Then errFile.Close
    Set errFS = Nothing
    Set errFile = Nothing
    Set wterms = Nothing
    Set xml = Nothing
    Set wNodes = Nothing
    Set curXmlWork = Nothing
End Sub


