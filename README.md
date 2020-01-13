These Microsoft Outlook Visual Basic MACROs export various notes or contacts into files formatted to use when importing into other programs or for backup purposes.  I have used this macros, or a derivation thereof, in Outlook 2007 through Outlook 2016.

Basically:

- Start Outlook and enable macros by accepting the security pop-up
- Create a new macro.  Copy & paste the VBScript code into the editor window.
- Execute the macro by clicking the Green right arrow in the editor or using the macro menu

There are multiple seperate macros in this repository:

# ExportOutlookContactsToVCF
Export Outlook Contacts Into Individual .VCF Files

It will:

-  Create an underlying directory structure:   C:\\OutlookContactsExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the CONTACTS folder to export
-  Create a individual .VCF files (one per contact) in the folder

# Export-OutlookNotes-To-LastPassCSV
Export Microsoft Outlook Notes Into LastPass Generic CSV Import Format

This is the initial release of the code after I used it to successfully export 426 Outlook notes into a single CSV file and then import them into LastPass as a series of Secure Notes.

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .CSV file suitable for importation into LastPass as a series of Secure Notes

# Export-OutlookNotes-To-XML
Export Microsoft Outlook Notes Into XML

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .XML file containing all note informaton
