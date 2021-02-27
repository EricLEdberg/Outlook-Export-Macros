These Microsoft Outlook Visual Basic MACROs export notes or contacts into files formatted to use when importing into other programs or for backup purposes.  I have used this macros, or a derivation thereof, in Outlook 2007 through Outlook 2016.

The macros are an easy read.

Basically:

- Start Outlook, type "macros" in the search box and enable macros by accepting the security pop-up
- Create a new macro.  Copy & paste the VBScript code into the editor window
- Execute the macro by clicking the Green right arrow in the macro editor or using the macro menu

There are several macros in this repository:

# Export-OutlookNotes-To-BitwardenCSV

Export Microsoft Outlook Notes In Bitwarden CSV Import Format

This macro will export all Notes in the selected folder into a single CSV file that can then be imported into Bitwarden as Secure Notes.  I routinely import 450+ notes into Bitwarden.

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the notes folder to export
-  Create a single .CSV file suitable for importation into Bitwarden as a series of Secure Notes
-  Bitwarden limits the length when importing secure notes to 1000 characters.  Outlook notes that exceed 1000 characters will be trucated.  A warning message will itentify truncated notes on completion.

# Export Microsoft Outlook Notes In LastPass Generic CSV Import Format

This macro will export all Notes in the selected folder into a single CSV file that can then be imported into LastPass as multiple Secure Notes.  I routinely import 450+ notes into LastPass.

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .CSV file suitable for importation into LastPass as a series of Secure Notes
# Outlook_Export_Notes_To_Text_Files
Export the selected Notes folder to individual text files

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create an individual files with the contents of every Outlook note

# ExportOutlookContactsToVCF
Export Outlook Contacts Into Individual .VCF Files

It will:

-  Create an underlying directory structure:   C:\\OutlookContactsExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the CONTACTS folder to export
-  Create an individual .VCF files (one per contact) in the folder

# Export-OutlookNotes-To-LastPassCSV
Export Microsoft Outlook Notes In LastPass Generic CSV Import Format

This macro will export all Notes in the selected folder into a single CSV file that can then be imported into LastPass as multiple Secure Notes.  I routinely import 450+ notes into LastPass.

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .CSV file suitable for importation into LastPass as a series of Secure Notes

# Export-OutlookNotes-To-XML
Export Microsoft Outlook Notes Into A Common XML

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .XML file containing all note informaton
