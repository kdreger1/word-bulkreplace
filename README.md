# word-bulkreplace
VBA module to do a bulk find-and-replace in Word using patterns defined in an Excel spreadsheet. Windows OS only.

Once you have imported the module, you will be able to use it in every Word document you use. To import the module:
1. Open up Word
2. Press Alt+F11 to open up the VBA Editor
3. In the editor, choose File -> Import File
4. Navigate to, and select, the PMAmodule_renumber_vN.bas file. (The N is the version number and may change over time.)

To use the module, you will first need to create an Excel file (.xlsx only). An example file called "testingInputPatterns.xlsx" is included in this repository. The first column are the words to search for; the second column are what each word is replaced by. Do not include column headers.

A "word" in the context of this module is a series of characters that form a discrete unit and is separated from surrounding text by a space or non-alphanumeric character (such as a period). For example, "101" will not match part of the string "101001" nor will it match "101a" but will match the string "101.". The only non-alphanumeric character that can be matched on is an underscore. Rows in the Excel spreadsheet that are missing a change-from field or a change-to field are ignored. Also ignored are rows with non-changes; example: 101a -> 101a.

There are three modes this find-and-replace feature can run in:
1. Highlight any changes, but do not make them
2. Highlight and make any changes
3. Make the changes without highlighting them

In any case, the changes are applied to a copy of the Word document that is created and becomes the active document. It is up to the user to name the new file and save it.

To trigger the module in the Word document you want to find-and-replace:
1. Click View on the ribbon, and then under Macros, choose View Macros. 
2. You will see the macro "PMA_renumber". Click on it (if it isn't already highlighted) and press the Run button.
3. A file selector window will appear. Navigate to, and select, the Excel spreadsheet with your patterns in it
4. A pop-up menu will appear. Select the mode you want to run in
5. Save the copy of the Word document that is created if you wish to keep it

It is possible to have multiple versions of PMA_renumber installed. However, it is recommended that you only keep the latest version. To delete modules:
1. Click View on the ribbon, and then under Macros, choose View Macros.
2. Select the module you want to delete and then press the Delete button
