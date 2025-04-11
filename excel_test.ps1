# Import the required modules
Import-Module -Name ImportExcel
$excelFilePath = "C:\Users\Michael.Barrett\Documents\testing\Excel test\testing_auto.xlsx"
#$workbook = New-ExcelWorkbook
#$worksheet = $workbook | Add-ExcelWorksheet -WorksheetName "Data"
#$worksheet | Write-ExcelRange -StartCell "A1" -Value $queryResult -AutoFit
#$workbook | Save-ExcelWorkbook -Path $excelFilePath
#$excelFilePath
#$workbook.Application.Quit()

#Create new empty excel file
"" | Export-Excel -Path $excelFilePath -WorksheetName "MyWorksheet"

#
# Write operations
#

#Open existing excel file
$ExcelPkg = Open-ExcelPackage -Path  $excelFilePath

#View or Modify excel file
$WorkSheet = $ExcelPkg.Workbook.Worksheets["sheet1"].Cells #open excel worksheet cells from worksheet "sheet1"

#Values can be accessed by row, column. Similar to a 2D array.
$WorkSheet[1,4].Value = "New Column Header" #Starts at index 1 not 0

#Load value at index
$ValueAtIndex = $WorkSheet[2,1].Value #Loads the value at row 2, column A

#The changes will not display in the Excel file until Close-ExcelPackage is called.  
Close-ExcelPackage $ExcelPkg #close and save changes made to the Excel file.

#If the file is currently in use, Close-ExcelPackage will return an error and will not save the information.

#
# Read operations
#

#Load the Excel file into a PSCustomObject
$ExcelFile = Import-Excel $excelFilePath -WorksheetName "Sheet1" 

#Load a column
$SpecificColumn = $ExcelFile."anotherHeader" #loads column with the header "anotherHeader" -- data stored in an array

#Load a row
$SpecificRow = $ExcelFile[1] #Loads row at index 1. Index 1 is the first row instead of 0.

# Map Contents to Hashtable to Interpret Data
# Sometimes mapping to a Hashtable is more convenient to have access to common Hashtable operations. Enumerate a Hashtable with the row's # data by:

$HashTable = @{}
$SpecificRow= $ExcelFile[2]
$SpecificRow.psobject.properties | ForEach-Object { 
    $HashTable[$_.Name] = $_.Value
}

#To then iterate through the enumerated Hashtable:

ForEach ($Key in ($HashTable.GetEnumerator()) | Where-Object {$_.Value -eq "x"}){ #Only grabs a key where the value is "x"
    #values accessible with $Key.Name or $Key.Value
}