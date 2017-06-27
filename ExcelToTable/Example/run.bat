rem Specified -outputfile, -worksheet, -format, -range
ExcelToTable.exe -filename dummydata.xlsx -outfile output_html.html -worksheet 1 -format html -range A4:C10

ExcelToTable.exe -filename dummydata.xlsx -outfile output_html.txt -worksheet 1 -format wikitable -range A4:C10 

ExcelToTable.exe -filename dummydata.xlsx -outfile output_html.jsonobjects.json -worksheet 1 -format jsonobjects -range A4:C10

ExcelToTable.exe -filename dummydata.xlsx -outfile output_html.jsonarrays.json -worksheet 1 -format jsonarrays -range A4:C10


rem Specified -outputfile, -worksheet, -format, defaults to all data.
ExcelToTable.exe -filename dummydata.xlsx -outfile output_html_alldata.html -worksheet 1 -format html

ExcelToTable.exe -filename dummydata.xlsx -outfile output_html_alldata.txt -worksheet 1 -format wikitable

ExcelToTable.exe -filename dummydata.xlsx -outfile output_html_alldata.jsonobjects.json -worksheet 1 -format jsonobjects

ExcelToTable.exe -filename dummydata.xlsx -outfile output_html_alldata.jsonarrays.json -worksheet 1 -format jsonarrays