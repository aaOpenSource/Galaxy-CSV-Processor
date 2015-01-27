Galaxy CSV Processor
=================

This code is meant to be used within an Excel VBA project to assist with the import and export of Galaxy Dump files from Wonderware System Platform.

##Installation
To "install" you should create a macro enabled workbook and import the three files.  There is a single form where the user can execute the functions.  Someone smarter than me can probably help put together a nice package or set of instructions for creating menus or buttons to run the functions.
 
## Functions
###Import
The import function allows the user to specify a single CSV file for import.  After import the code will automatically apply the built in text to columns step.  After splitting the data into columns the code will create a new worksheet for each different template type and move the appropriate contents to that sheet.

##Save
The save functions works in reverse of the import function.  The save function will take some or all of the worksheets and combine them into a single worksheet.  After combining the worksheets the code will execute a save as CSV to export a single CSV file. The user may select to export all sheets or just selected sheets.

## Contributors
* [Andy Robinson](mailto:andy@phase2automation.com), Principal of [Phase 2 Automation](http://phase2automation.com).
* Billy Rios of Laconicly

## License

MIT License. See the [LICENSE file](/LICENSE) for details.