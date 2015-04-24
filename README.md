Galaxy CSV Processor
=================

This code is meant to be used within an Excel VBA project to assist with the import and export of Galaxy Dump files from Wonderware System Platform.

##Installation
To "install" you should create a macro enabled workbook and import the three files.  There is a single form where the user can execute the functions.  

Also due to custom export function you will need to add a reference to Microsoft ActiveX Data Objects 2.5 Library under Tool->References when yoou are in VBA editor.

More Detailed Instructions: 
 To import the files, press Alt+F11 to open the VBA editor from Excel.  Go to File --> Import and browse to the file location.  Note that when you import the .FRM file, the .FRX file is imported automatically.  

 To actually use any of this, you will need to run the macro that you just imported so you can access the form that you just imported.  Save the file and exit the VBA editor to return to the Excel worksheet.  Go to the "Insert" menu in the ribbon at the top, and select "Shapes".  Pick any shape (I personally prefer the smiley face) and customize it to look like whatever you want.  Right click on the shape and pick "Assign Macro" from the context menu.  Select the "ShowForm" macro and click "OK".  Save the file as an Excel Macro Enabled Work Book (.xlsm).

## Functions
###Import
The import function allows the user to specify a single CSV file for import.  After import the code will automatically apply the built in text to columns step.  After splitting the data into columns the code will create a new worksheet for each different template type and move the appropriate contents to that sheet.

##Save
The save functions works in reverse of the import function.  The save function will take some or all of the worksheets and combine them into a single worksheet.  After combining the worksheets the code will execute a save as CSV to export a single CSV file. The user may select to export all sheets or just selected sheets.  One key point to consider is that this code will synthesize CSV text as opposed to automatically generating from Excel.  The reason is that Excel can not directly save CSV as UTF-8.  So we create CSV text and save as UTF-8.  This was driven off a user request and filed as [Issue #1](/../../issues/1)

## Contributors
* [Andy Robinson](mailto:andy@phase2automation.com), Principal of [Phase 2 Automation](http://phase2automation.com).
* Anonymous - a special thanks to an anonymous contributor who provided code and inspiration to finish the work that I started long ago.  Thanks!

## License

MIT License. See the [LICENSE file](/LICENSE) for details.
