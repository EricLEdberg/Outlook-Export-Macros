# Export-OutlookNotes-To-LastPassCSV
Export Microsoft Outlook Notes Into LastPass Generic CSV Import Format

This is a Microsoft Outlook Visual Basic MACRO that exports the selected Outlook "Notes" folder into a CVS data file formatted to use when importing (creating) multiple LastPass Secure Notes.  I have used this macro, or a derivation thereof, in Outlook 2007 through Outlook 2016.

This is the initial release of the code after I used it to successfully export 426 Outlook notes into a CSV file and and import them into my LastPass account.

Eventually, I may update the instructions to provide enhanced instructions.

Basically:

- Start Outlook and enable macros by accepting the security pop-up
- Create a new macro.  Copy & paste the VBScript code into the editor window.
- Execute the macro by clicking the Green right arrow in the editor or using the macro menu

It will:

-  Create an underlying directory structure:   C:\OutlookNotes\Export\%COMPUTERNAME%\%MMDDYY%
-  Prompt you to select the notes folder to export
-  Create a single .CSV file in the folder suitable for importation into LastPass as a series of Secure Notes
