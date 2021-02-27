'# -----------------------------------------------------------------------
'# Export Outlook Notes Into Bitwarden CSV Import Format (should really use JSON import)
'# https://bitwarden.com/help/article/condition-bitwarden-import/
'# Create a UTF-8 encoded plaintext file with the following header as the first line in the file:
'#
'# folder,favorite,type,name,notes,fields,login_uri,login_username,login_password,login_totp
'# Social,1,login,Twitter,,,twitter.com,me@example.com,password123,
'# ,,login,EVGA,,,https://www.evga.com/support/login.asp,hello@bitwarden.com,fakepassword,TOTPSEED123
'# ,,login,My Bank,Bank PIN is 1234,"PIN: 1234",https://www.wellsfargo.com/home.jhtml,john.smith,password123456,
'# ,,note,My Note,"This is a secure note.",,,,,
'#
'#
'# Author:   Eric L. Edberg   2/2021
'#
'# -----------------------------------------------------------------------

Public Sub ExportNotesToBitwardenCSV()

    DoCheck = False
    GenericHeader = "folder,favorite,type,name,notes,fields,login_uri,login_username,login_password,login_totp"
    

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
    
    MsgBox "Exporting Notes in Bitwarden CSV format to: " & ExportFolder
    
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        MkDir ExportFolder
    ElseIf DoCheck = True Then
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
    
    '# for LastPass, all notes save to common CSV file
    xNoteFile = ExportFolder & "\Bitwarden_SecureNotesImport_" & YYYYMMDD & ".csv"
    
    '# TODO:  should generate random file postfix here
    If (fso.FileExists(xNoteFile)) Then
        xRI = (Int((max - min + 1) * Rnd + min))
        xNoteFile = ExportFolder & "\Bitwarden_SecureNotesImport_" & YYYYMMDD & xRI & ".csv"
        MsgBox "Import file already exists.  Saving to random import file: " & xNoteFile
    End If
    
    Open xNoteFile For Output As 1
    
    '# -----------------------------------------------------------------------
    '# Write generic cvs header as first line
    '# -----------------------------------------------------------------------
    Print #1, GenericHeader
        
    '# -----------------------------------------------------------------------
    '# -----------------------------------------------------------------------
    xMessage = ""
    For cnt = 1 To myNote.Items.Count
    
        '# Subject is 1st line of the note Body
        xSubject = Replace(Replace(Replace(myNote.Items(cnt).Subject, "/", "-"), "\", "-"), ":", "-")
        xSubject = Trim(xSubject)
        
        '# Double Quote Quote character
        xBody = myNote.Items(cnt).Body
        xBody = Replace(xBody, Chr(34), Chr(34) & Chr(34))
        
        '# Bitwarden only supports 1000 characters during import
        xLen = Len(xBody)
        If (xLen > 1000) Then
            xMessage = xMessage & vbCrLf & "TOLONG(" & xLen & ") TRUNCGATING to 1000 Chars: " & xSubject
            xBody = Left(xBody, 1000)
        End If
        
             
        '# Secure Notes import CSV layout from Bitwarden
        '# folder,favorite,type,name,notes,fields,login_uri,login_username,login_password,login_totp
        '# ,,note,My Note,"This is a secure note.",,,,,
        
        xFolder = "Outlook Notes" & ","
        xFavorite = ","
        xType = "note,"
        xName = xSubject & ","
        xNote = """" & xBody & """" & ","
        xExtra = ",,,,"
        
        Print #1, xFolder; xFavorite; xType; xName; xNote; xExtra; "0"
                                    
        
'        If cnt = 3 Then
'            Close #1
'            MsgBox "Successfully Exported: " & cnt & ", notes into Bitwarden import CSV file: " & xNoteFile
'            Close #1
'            Exit Sub
'        End If

    Next
     
    Close #1
    
    If (Not IsEmpty(xMessage)) Then
        MsgBox xMessage
    End If
    
    
    MsgBox "Successfully Exported: " & cnt & ", notes into LastPass import CSV file: " & xNoteFile
    
End Sub
