'# -----------------------------------------------------------------------
'# https://technet.microsoft.com/en-us/magazine/2008.02.heyscriptingguy.aspx?pr=PuzzleAnswer
'# -----------------------------------------------------------------------

Public Sub ExportNotesToXML()
    Dim ExportFolder As String
    Dim YYYYMMDD As String
    Dim strHostName As String
    Dim strIPAddress As String
        
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    Set objRoot = xmlDoc.createElement("OutlookNotes")
    xmlDoc.appendChild objRoot
        
    ExportFolder = "C:\OutlookNotesExport"
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        MkDir ExportFolder
    End If
    
    strHostName = Environ$("computername")
    ExportFolder = "C:\OutlookNotesExport\" & strHostName
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        MkDir ExportFolder
    End If
        
    YYYYMMDD = Format(Date, "yyyymmdd")
    ExportFolder = "C:\OutlookNotesExport\" & strHostName & "\" & YYYYMMDD
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
       MkDir ExportFolder
    End If
     
    MsgBox "Exporting Notes as 1 XML file in folder: " & ExportFolder
       
    Set myNote = Application.GetNamespace("MAPI").PickFolder
    For cnt = 1 To myNote.Items.Count
                 
        Set objRecord = xmlDoc.createElement("Note")
        objRoot.appendChild objRecord
         
        '# SUBJECT
        xSubject = Replace(Replace(Replace(myNote.Items(cnt).Subject, "/", "-"), "\", "-"), ":", "-")
        Set objSubject = xmlDoc.createElement("Subject")
        objSubject.Text = xSubject
        objRecord.appendChild objSubject
                
        '# Last Modification Time
        Set objLastModificationTime = xmlDoc.createElement("LastModificationTime")
        objLastModificationTime.Text = myNote.Items(cnt).LastModificationTime
        objRecord.appendChild objLastModificationTime
                
        '# Note Body
        Set objBody = xmlDoc.createElement("Body")
        objBody.Text = myNote.Items(cnt).Body
        objRecord.appendChild objBody
        
        '# Save individual notes to a file
        '# myNote.Items(cnt).SaveAs ExportFolder & "\" & xSubject & ".txt", OlSaveAsType.olTXT
     Next
     
    '# Save XML document to File
    Set objIntro = xmlDoc.createProcessingInstruction("xml", "version='1.0'")
    xmlDoc.InsertBefore objIntro, xmlDoc.ChildNodes(0)
    xmlDoc.Save ExportFolder & "\OutlookNotes.xml"
          
    MsgBox "Completed Exporting Notes as 1 XML files into folder: " & ExportFolder
    
End Sub