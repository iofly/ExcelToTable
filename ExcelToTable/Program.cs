using System;
using System.Collections.Generic;

namespace ExcelToTable
{
    public enum ResultCode { Success, ErrorOpeningFile };

    class Program
	{
        static void Main(string[] args)
        {
            ResultCode rc;
            string ResultDesc = string.Empty;
            List<List<string>> rows;
            string fileext = ".txt";
            string text = string.Empty;

            #region Define supported arguments
            List<SimpleArg> supportedArgs = new List<SimpleArg>();
            supportedArgs.Add(new SimpleArg { Name = "-filename", IsSwitch = false, Required = true, DefaultValue = null, ArgType = SimpleArgType.ExistingFilename, ExmaplePlaceholder="excelfilename", Description="Required. The Microsoft Excel file name" });
            supportedArgs.Add(new SimpleArg { Name = "-outfile", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String, ExmaplePlaceholder = "outputfilename", Description = "Optional. Output file. Defaults to [excelfilename] with format specific extension appended." });
            supportedArgs.Add(new SimpleArg { Name = "-format", IsSwitch = false, Required = false, DefaultValue = "html", ArgType = SimpleArgType.String, ExmaplePlaceholder = "html|wikitable|jsonsobjects|jsonarrays", Description = "Optional. Output file format [html|wikitable|jsonsobjects|jsonarrays]. Defaults to html." });
            supportedArgs.Add(new SimpleArg { Name = "-worksheet", IsSwitch = false, Required = false, DefaultValue = 1, ArgType = SimpleArgType.Integer, ExmaplePlaceholder = "1-n", Description = "Optional. A one-based index of the worksheet to export data from. Defaults to 1." });
            supportedArgs.Add(new SimpleArg { Name = "-range", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String, ExmaplePlaceholder = "excelrange", Description = "Optional. Excel cell range to export. e.g. A12:C23. Defaults to the worksheet's used extents." });
            #endregion

            #region Parse and validate arguments
            
            SimpleArgParser parser = new SimpleArgParser(supportedArgs);
            Dictionary<string, dynamic> ar = new Dictionary<string, dynamic>();

            //Validate against general rules as specified above.
            try
            {
                ar = parser.ParseArgs(args);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                parser.ShowUsage();
                return;
            }

            //Fix for filename ->abs filename - for case where app is called via batch file
            if (!System.IO.Path.IsPathRooted(ar["-filename"]))
            {
                ar["-filename"] = System.IO.Path.GetFullPath(ar["-filename"]);
            }


            string outfile = string.Empty;
            if(ar.ContainsKey("-outfile"))
            {
                outfile = System.IO.Path.GetFullPath(ar["-outfile"]);
            }

            WorkSheetRangeCoordinates wsrc = null;
            if (ar.ContainsKey("-range"))
            {
                wsrc = ExcelReader.ParseExcelRange(ar["-range"]);
            }

            //App specific validation of args
            try
            {
                Utils.ValidateAppArgs(ar);
            }
            catch(ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                parser.ShowUsage();
                return;
            }
            #endregion

            #region Read excel file
            rows = ExcelReader.ReadExcelRows(ar["-filename"], out rc, out ResultDesc, ar["-worksheet"], wsrc);
            if (rc == ResultCode.ErrorOpeningFile)
            {
                Console.WriteLine(String.Format("ErrorOpeningFile: {0}", ResultDesc));
                return;
            }
            #endregion

            #region Output result
            if (ar["-format"] == "wikitable")
            {
                text = Utils.RowsToWikiTable(rows);
                fileext = ".txt";
            }
            else if (ar["-format"] == "jsonobjects")
            {
                text = Utils.RowsToJSON_ArrayOfObjects(rows);
                fileext = ".objects.json";
            }
            else if (ar["-format"] == "jsonarrays")
            {
                text = Utils.RowsToJSON_ArrayOfArrays(rows);
                fileext = ".arrays.json";
            }
            else
            {
                text = Utils.RowsToHTMLTable(rows);
                fileext = ".html";
            }


            if (!String.IsNullOrWhiteSpace(outfile))
            {
                System.IO.File.WriteAllText(outfile, text);
            }
            else
            {
                System.IO.File.WriteAllText(System.IO.Path.GetFileName(ar["-filename"]) + fileext, text);
            }
            #endregion
        }
    }
}
