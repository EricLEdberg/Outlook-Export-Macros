'# -----------------------------------------------------------------------
'# https://docs.microsoft.com/en-us/office/vba/api/outlook.contactitem
'# Author:  Eric L. Edberg 2019-10-31
'# Select and export Outlook Contacts Folder to .VCF files
'# -----------------------------------------------------------------------

Public Sub ExportContactsToVCF()
    Dim ExportFolder As String
    Dim YYYYMMDD As String
    Dim strHostName As String
    Dim objFSO
    
    Const max = 1000
    Const min = 1
    Const olFolderContacts = 10
    Const olVCard = 6
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Randomize
    
    ExportFolder = "C:\OutlookContactsExport"
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        MkDir ExportFolder
    End If
    
    strHostName = Environ$("computername")
    ExportFolder = "C:\OutlookContactsExport\" & strHostName
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        MkDir ExportFolder
    End If
        
    YYYYMMDD = Format(Date, "yyyymmdd")
    ExportFolder = "C:\OutlookContactsExport\" & strHostName & "\" & YYYYMMDD
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
       MkDir ExportFolder
    End If
        
    MsgBox "Exporting Outlook Contacts as .VCF files into folder: " & ExportFolder
    
    MsgBox "During the next pop-up prompt, select the Outlook Contacts folder that you wish to export as .VCF contacts"
    Set objContacts = Application.GetNamespace("MAPI").PickFolder
    
    MsgBox "The next step will export your contacts.  It may take awhile and Outlook will appear to hang (lock up) during this process.  It may take a minute or 3 to complete depending on the number of contacts you have to export"
        
    On Error Resume Next
    
    ErrorCnt = 0
    OkCnt = 0
    Err = 0
    DupCnt = 0
    For cnt = 1 To objContacts.Items.Count
        strName2 = objContacts.Items(cnt).LastName & objContacts.Items(cnt).FirstName
        strName = objContacts.Items(cnt).FileAs
        If strName = "" Or IsNull(strName) Then
            strName = objContacts.Items(cnt).CompanyName
        End If

		'# ToDo:  Is there a function to make a string filename safe?
        strName = Replace(Replace(Replace(strName, "(", "-"), ")", "-"), "*", "-")
        strName = Replace(Replace(Replace(strName, "&", "-"), " ", "-"), "*", "-")
        strName = Replace(Replace(Replace(strName, "/", "-"), "\", "-"), ":", "-")
        
        strPath = ExportFolder & "\" & strName & ".vcf"
           
        '# 
        If (fso.FileExists(strPath)) Then
            xRI = (Int((max - min + 1) * Rnd + min))
            strPath = ExportFolder & "\" & strName & "-" & xRI & ".vcf"
            DupCnt = DupCnt + 1
            '# MsgBox "Contact already exported:  Saving duplicate with random postfix: " & strName & "-" & xRI & ".vcf"
        End If
           
        Err = 0
        objContacts.Items(cnt).SaveAs strPath, olVCard
    
        If Err Then
            ErrorCnt = ErrorCnt + 1
        Else
            OkCnt = OkCnt + 1
        End If
        
    Next
    
    On Error GoTo 0
    
    MsgBox "COMPLETED exported (" & OkCnt & ") contacts, with (" & DupCnt & ") duplicates, with (" & ErrorCnt & ") errors, folder: " & ExportFolder
    
End Sub
