
'# -----------------------------------------------------------------------
'# http://www.gregthatcher.com/Scripts/VBA/Outlook/GetListOfNotes.aspx
'# Greatly enhanced by Eric L. Edberg
'# -----------------------------------------------------------------------

Public Sub ExportNotesToText()
    Dim ExportFolder As String
    Dim YYYYMMDD As String
    Dim strHostName As String
    Dim strIPAddress As String
    Dim objFSO
    
    Set fso = CreateObject("Scripting.FileSystemObject")
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
    
    MsgBox "Exporting Notes To: " & ExportFolder
      
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        MkDir ExportFolder
    Else
        MsgBox "Oops, the export folder already exists." & vbCrLf & "Will not overwrite previous export files." & vbCrLf & "Please rename if necessary." & vbCrLf & "Folder: " & ExportFolder
        Exit Sub
    End If
       
    MsgBox "During the next prompt, select the Notes folder that you want to export..."
    
    Set myNote = Application.GetNamespace("MAPI").PickFolder
     
    '# TODO:  validate that a Notes folder was really selected...
    
    MsgBox "About to export: " & myNote.Items.Count & ", notes.  This may take a while..."
    
    '# random integer between 1 - 1000
    Dim max, min
    max = 1000
    min = 1
    Randomize
    
    For cnt = 1 To myNote.Items.Count
    
        '# Subject is 1st line of the note Body
        xSubject = Replace(Replace(Replace(myNote.Items(cnt).Subject, "/", "-"), "\", "-"), ":", "-")
        xSubject = Trim(xSubject)
        xNote = myNote.Items(cnt).Body
        xNoteFile = ExportFolder & "\" & xSubject & ".txt"
        
        '# TODO:  should generate random file postfix here
        If (fso.FileExists(xNoteFile)) Then
            xRI = (Int((max - min + 1) * Rnd + min))
            xNoteFile = ExportFolder & "\" & xSubject & "-" & xRI & ".txt"
            MsgBox "Note already exists:  Saving duplicate note with random postfix e.g.: " & xSubject & "-" & xRI & ".txt"
        End If
             
        '# TODO:  obtain Modified Date of Note and set the date of the file to the same date value
             
        
        Open xNoteFile For Output As 1
        Print #1, xNote
        Close #1
                 
        '# SaveAs includes Modified header which we don't want
        '# myNote.Items(cnt).SaveAs ExportFolder & "\" & xSubject & ".txt", OlSaveAsType.olTXT
        
    Next
     
    MsgBox "Successfully Exported: " & myNote.Items.Count & ", Notes Into Folder: " & ExportFolder
End Sub

