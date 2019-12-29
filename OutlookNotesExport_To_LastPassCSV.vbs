'# -----------------------------------------------------------------------
'# Author:  Eric L. Edberg - 12/2019
'# See:     https://support.logmeininc.com/lastpass/help/how-do-i-import-stored-data-into-lastpass-using-a-generic-csv-file
'# -----------------------------------------------------------------------

Public Sub ExportNotesToLastPassSecureNoteCSV()

    DoCheck = True
    GenericHeader = "url,username,password,extra,name,grouping,fav" 

    '# -----------------------------------------------------------------------
    '# this section is common and should be a common function
    '# -----------------------------------------------------------------------
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
    
    MsgBox "Exporting Notes in LastPass format to: " & ExportFolder
    
    If DoCheck = True Then
        If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
            MkDir ExportFolder
        Else
            MsgBox "Oops, the export folder already exists." & vbCrLf & "Will not overwrite previous export files." & vbCrLf & "Please rename if necessary." & vbCrLf & "Folder: " & ExportFolder
            Exit Sub
        End If
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
    
    '# for LastPass, all notes save to common CSV file
    xNoteFile = ExportFolder & "\LastPass_SecureNotesImport_" & YYYYMMDD & ".csv"
    
    '# TODO:  should generate random file postfix here
    If (fso.FileExists(xNoteFile)) Then
        xRI = (Int((max - min + 1) * Rnd + min))
        xNoteFile = ExportFolder & "\LastPass_SecureNotesImport_" & YYYYMMDD & xRI & ".csv"
        MsgBox "Import file already exists.  Saving to random import file: " & xNoteFile
    End If
    
    Open xNoteFile For Output As 1
    
    '# -----------------------------------------------------------------------
    '# Write generic cvs header as first line
    '# -----------------------------------------------------------------------
    Print #1, GenericHeader
        
    '# -----------------------------------------------------------------------
    '# -----------------------------------------------------------------------
    For cnt = 1 To myNote.Items.Count
    
        '# Subject is 1st line of the note Body
        xSubject = Replace(Replace(Replace(myNote.Items(cnt).Subject, "/", "-"), "\", "-"), ":", "-")
        xSubject = Trim(xSubject)
        
        '# Double Quote Quote character
        xBody = myNote.Items(cnt).Body
        xBody = Replace(xBody, Chr(34), Chr(34) & Chr(34))
             
        '# Secure Notes import CSV layout from LastPass
        '# URL , UserName, Password, extra, Name, grouping, fav
        '# http://sn,,,note content,Secure Notes,0
        xURL = "http://sn" & ","
        xUserName = ","
        xPassword = ","
        xExtra = """" & xBody & """" & ","
        xName = xSubject & ","
        xFolder = "Outlook Notes" & ","
        
        Print #1, xURL; xUserName; xPassword; xExtra; xName; xFolder; "0"

    Next
     
    Close #1
    
    MsgBox "Successfully Exported: " & cnt & ", notes into LastPass import CSV file: " & xNoteFile
    
End Sub
