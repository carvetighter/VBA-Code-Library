# VBA-Code-Library

The VBA library is a set of methods that I have used over the years to help me work accurately and efficiently.  I inldued the modules so others can use and modify for their own use.

The Excel file has all the modules and all the references included to be able to use once downloaded.

Mod_Autofilter.bas methods:
  - AutoFilter_Clear -> clears the autofilter
  - Autofilter_OnOff -> turns the autogilter on or off
  - Autofilter_SimpleSort -> simple sort

Mod_Cells.bas methods:
  - Cells_Format -> formats all the cells in a worksheet
  - Cells_IsString -> determines if a cell values is a string
  - Cells_ReturnNumberOrLetters -> used in Row_GetLast and Column_GetLast for the return value but not limited to only those methods

Mod_Charts.bas methods:
  - Chart_New -> creates a graph with a single series based on a specific format
  - Chart_Line -> creates a line graph based on a simple format
  - Chart_DeleteAll -> detltes all the graphs in a worksheet

Mod_Column.bas methods:
  - Column_CopyPaste -> copy and paste a column from one worksheet to another
  - Column_Find -> find a column based on a string passed
  - Column_FindByDate -> find a column based on a date
  - Column_FindByTimePeriod -> find a column based on the time period from the last column with a date
  - Column_GetCriteria -> get all the unique values in a column
  - Column_GetCriteriaByColumn -> get all the unique values of the columsn from the first column to the last column with data in it
  - Column_GetLast -> get the last column in a row with data in it
  - Column_InsertWithHeader -> insert a column with a header to the left of a column
  - Column_TakeOutBlankCells -> delete all the rows with no values in a column

Mod_Date.bas methods:
  - Date_GetHoursMinutesSeconds -> finds the hours, minutes and seconds

Mod_Files.bas methods:
  - Files_ArchiveSource -> archives all the files in the source folder to an archive folder
  - Files_Count -> count the number of files in a folder
  - Files_FindXlsx -> find *.xlsx files for the OSx version of Microsoft Office
  - Files_GetNames -> get the names of all files in a folder
  - Files_GetPath -> get the file path that user selects from a dialog window

Mod_Main.bas methods:
  - Main -> template for a main method

Mod_Outlook.bas methods:
  - Application_ItemSend -> BCC every email sent to a specified email
  - Outlook_SendEmail -> example of sending an email from Excel through Outlook

Mod_Pivot.bas methods:
  - Pivot_Get -> finds and returns specifid pivot table
  - Pivto_SetFieldValues -> set specific values in a pivot field

Mod_PowerPoint methods:
  - PowerPoint_AppendsixSlide -> an example on how to create a PowerPoint slide from a PowerPoint presentation and add data / objects         from Excel
  - PowerPoint_CutPasteChartsFromExcel -> an example of how to cut and paste charts from Excel to PowerPoint
  - PowerPoint_IDShape -> identifies the type of PowerPoint shape
  - PowerPoint_ModifyTitle -> an example of how to modify a slide title in PowerPoint
  - PowerPoint_TestGroupItems -> determines if a shape is grouped or not

Mod_Queries.bas methods:
  - Query_Refresh -> a simple method that refreshes the SQL queries of all the workshees in the workbooks collection
  - Query_Website -> copies all the information from a website to an Excel worksheet

Mod_Row.bas methods:
  - Row_Copy -> copies row from one worksheet to another
  - Row_CountVisible -> counts the visible cells in a row
  - Row_Find -> finds a row based on a specific criteria
  - Row_FindAndDelete -> finds a row and deletes it
  - Row_FindByFirstWord -> finds a row by the first work in the cell
  - Row_FindTransition -> finds the row of a column where it transitoins from data in a series of cells to no data in a cell or no data                             in a series of cells to data in a cell
  - Row_GetCriteria -> gets all the unique values of a row
  - Row_GetCriteriaByRow -> gets all the unique values of the rows from the first row to the last row ith data in it
  - Row_GetLast -> finds the last row with data in a specified column
  - Row_Insert -> insert a specifiec number of rows above a specific row
  - Row_Unhide -> unhides all rows between two specified rows

Mod_Search.bas methods:
  - Search_Log -> conducts a simple log search (not optimized) of column

Mod_String.bas methods:
  - String_FindAllPositOfChar -> finds all the positions of a single charcter in a string

Mod_Workbook.bas methods:
  - Workbook_Clear -> formats every worksheet in a workbook
  - Workbook_Find -> finds a specific workbook in the Excel application by name
  - Workbook_FindOrCreate -> finds a specific workbook and if not creates a new one

Mod_Worksheet.bas methods:
  - Worksheet_Clear -> formats a worksheet to a specified format
  - Worksheet_FindOrCreate -> finds a worksheet by the name or if not creates a new one
