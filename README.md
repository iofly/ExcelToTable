# ExcelToTable

Tool written in C# to export text data from Excel to html, JSON or wikitable format.

Example usage
```
ExcelToTable.exe -filename dummydata.xlsx -outfile output.html -worksheet 1 -format html -range A4:C10
```

Usage as displayed by application
```
Usage: ExcelToTable.exe -filename excelfilename -outfile [outputfilename] -format [html|wikitable|jsonsobjects|jsonarrays] -worksheet [1-n] -range [excelrange]


-filename:              Required. The Microsoft Excel file name
-outfile:               Optional. Output file. Defaults to [excelfilename] with format specific extension appended.
-format:                Optional. Output file format [html|wikitable|jsonsobjects|jsonarrays]. Defaults to html.
-worksheet:             Optional. A one-based index of the worksheet to export data from. Defaults to 1.
-range:         Optional. Excel cell range to export. e.g. A12:C23. Defaults to the worksheet's used extents.
```
