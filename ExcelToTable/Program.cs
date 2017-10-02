using System;
using System.Collections.Generic;

namespace ExcelToTable
{
	public enum ResultCode { Success, ErrorOpeningFile };

	class Program
	{
		static void Main(string[] args)
		{
			//Defines supported arguments
			var supportedArgs = new List<SimpleArg>
			{
				new SimpleArg
				{
					Name = "-filename",
					IsSwitch = false,
					Required = true,
					DefaultValue = null,
					ArgType = SimpleArgType.ExistingFilename,
                    ExampleValuePlaceholder = "excelfilename",
					Description = "Required. The Microsoft Excel file name"
				},
				new SimpleArg
				{
					Name = "-outfile",
					IsSwitch = false,
					Required = false,
					DefaultValue = null,
					ArgType = SimpleArgType.String,
                    ExampleValuePlaceholder = "outputfilename",
					Description = "Optional. Output file. Defaults to [excelfilename] with format specific extension appended."
				},
				new SimpleArg
				{
					Name = "-format",
					IsSwitch = false,
					Required = false,
					DefaultValue = "html",
					ArgType = SimpleArgType.ValueRange,
                    ExampleValuePlaceholder = "html|wikitable|jsonsobjects|jsonarrays|excel",
					Description = "Optional. Output file format [html|wikitable|jsonsobjects|jsonarrays|excel]. Defaults to html.",
					ValueRange = {"html", "wikitable", "jsonsobjects", "jsonarrays", "excel"}
				},
				new SimpleArg
				{
					Name = "-worksheet",
					IsSwitch = false,
					Required = false,
					DefaultValue = 1,
					ArgType = SimpleArgType.Integer,
                    ExampleValuePlaceholder = "1-n",
					Description = "Optional. A one-based index of the worksheet to export data from. Defaults to 1."
				},
				new SimpleArg
				{
					Name = "-range",
					IsSwitch = false,
					Required = false,
					DefaultValue = null,
					ArgType = SimpleArgType.ExcelRange,
                    ExampleValuePlaceholder = "excelrange",
					Description = "Optional. Excel cell range to export. e.g. A12:C23. Defaults to the worksheet's used extents."
				}
            };
			SimpleArgParser parser;
			try
			{
				//Parse and validate arguments
				parser = Utils.GetArguments(supportedArgs, args);

				if (parser == null)
				{
					Console.WriteLine(String.Empty);
					SimpleArgParser.ShowUsage(supportedArgs);
                    return;
				}
			}
			catch(Exception ex)
			{
				Console.WriteLine(String.Empty);
				Console.WriteLine($"Error: {ex.Message}");
				SimpleArgParser.ShowUsage(supportedArgs);
                return;
			}

			//Read excel file
			var rows = ExcelAutomate.ReadExcelRows(parser.ParsedArguments["-filename"], 
											out ResultCode rc, 
											out string resultDesc, 
											parser.ParsedArguments["-worksheet"],
											parser.ParsedArguments.ContainsKey("-range") ? parser.ParsedArguments["-range"] : null);
			if (rc == ResultCode.ErrorOpeningFile)
			{
				Console.WriteLine($"ErrorOpeningFile: {resultDesc}");
				return;
			}

			//Produce output
			Utils.GenerateOutputFile(parser.ParsedArguments["-format"], 
											rows, 
											parser.ParsedArguments.ContainsKey("-outfile") ? parser.ParsedArguments["-outfile"] : null);
		    
		}
	}
}
