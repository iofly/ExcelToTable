using System;
using System.Collections.Generic;
using SimpleArgs;

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
            string text = string.Empty;


            //Define supported arguments
            var supportedArgs = new List<SimpleArg>();
            supportedArgs.Add(new SimpleArg { Name = "-filename", IsSwitch = false, Required = true, DefaultValue = null, ArgType = SimpleArgType.ExistingFilename, ExmaplePlaceholder="excelfilename", Description="Required. The Microsoft Excel file name" });
            supportedArgs.Add(new SimpleArg { Name = "-outfile", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String, ExmaplePlaceholder = "outputfilename", Description = "Optional. Output file. Defaults to [excelfilename] with format specific extension appended."});
            supportedArgs.Add(new SimpleArg { Name = "-format", IsSwitch = false, Required = false, DefaultValue = "html", ArgType = SimpleArgType.String, ExmaplePlaceholder = "html|wikitable|jsonsobjects|jsonarrays", Description = "Optional. Output file format [html|wikitable|jsonsobjects|jsonarrays]. Defaults to html." });
            supportedArgs.Add(new SimpleArg { Name = "-worksheet", IsSwitch = false, Required = false, DefaultValue = 1, ArgType = SimpleArgType.Integer, ExmaplePlaceholder = "1-n", Description = "Optional. A one-based index of the worksheet to export data from. Defaults to 1." });
            supportedArgs.Add(new SimpleArg { Name = "-range", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.ExcelRange, ExmaplePlaceholder = "excelrange", Description = "Optional. Excel cell range to export. e.g. A12:C23. Defaults to the worksheet's used extents." });
            
            
            //Parse and validate arguments
            SimpleArgParser parser = Utils.GetArguments(supportedArgs, args);
            if (parser == null) return;


            //Read excel file
            rows = ExcelReader.ReadExcelRows(parser.ParsedArguments["-filename"], out rc, out ResultDesc, parser.ParsedArguments["-worksheet"], parser.ParsedArguments["-range"]);
            if (rc == ResultCode.ErrorOpeningFile)
            {
                Console.WriteLine(String.Format("ErrorOpeningFile: {0}", ResultDesc));
                return;
            }

            //Produce output
            Utils.GenerateOutputFile(parser.ParsedArguments["-format"], rows, parser.ParsedArguments["-filename"], parser.ParsedArguments["-outfile"]);
        }
    }
}
