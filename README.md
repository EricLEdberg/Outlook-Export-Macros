These Microsoft Outlook Visual Basic MACROs export notes or contacts into files formatted to use when importing into other programs or for backup purposes.  I have used this macros, or a derivation thereof, in Outlook 2007 through Outlook 2016.

The macros are an easy read.

Basically:

- Start Outlook, type "macros" in the search box and enable macros by accepting the security pop-up
- Type "macros" into search again and create a new macro
- Copy & paste the VBScript code into the editor window
- Execute the macro by clicking the Green right arrow in the macro editor or using the macro menu

There are several macros in this repository:

# Export A Microsoft Outlook Notes Folder Into Bitwarden CSV Import Format

This macro will export all Outlook Notes in the selected folder into a single CSV file that can be imported into Bitwarden as multiple Secure Notes.

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%, to store export data
-  Prompt you to select the Outlook notes folder to export
-  Prompt for the name of the Bitwarden folder where imported secure notes will reside.  It is wise to input a new unique folder name and manually move the notes into other folders.  Note that the Bitwarden import process will not overwrite secure notes with the same name and blindly create duplicate items with the same name.
-  Create a single .CSV file suitable for importation into Bitwarden as a series of Secure Notes
-  Replace accented characters which are not supported by Bitwarden import
-  Bitwarden limits Secure Note import length to 10000 characters.  When creating multi-page notes, it will split an Outlook Note into 9999 character chunks and create seperate Bitwarden Secure Notes using a page counter appended to the note name.

# Export Microsoft Outlook Notes In LastPass Generic CSV Import Format

This macro will export all Notes in the selected folder into a single CSV file that can be imported into LastPass as multiple Secure Notes.

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .CSV file suitable for importation into LastPass as a series of Secure Notes

# Export Microsoft Outlook Notes Into Individual Text Files

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create an individual files with the contents of every Outlook note

# Export Microsoft Outlook Contacts Into Individual VCF Text Files

It will:

-  Create an underlying directory structure:   C:\\OutlookContactsExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the CONTACTS folder to export
-  Create an individual .VCF files (one per contact) in the folder

# Export Microsoft Outlook Notes Into A Single XML File

It will:

-  Create an underlying directory structure:   C:\\OutlookNotesExport\\%COMPUTERNAME%\\%MMDDYY%
-  Prompt you to select the NOTES folder to export
-  Create a single .XML file containing all note informaton
