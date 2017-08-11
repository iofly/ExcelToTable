using System;
using System.Collections.Generic;
using System.Text;
using System.Web;
using System.IO;

namespace ExcelToTable
{
	public static class Utils
	{
		public static string RowsToWikiTable(List<List<string>> rows)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < rows.Count; i++)
			{
				if (i == 0)
				{
					sb.Append("{| class=\"wikitable\"");
					sb.Append(Environment.NewLine);
					sb.Append("|-");
					sb.Append(Environment.NewLine);
					sb.Append("! ");
					for (int n = 0; n < rows[i].Count; n++)
					{
						sb.Append(rows[i][n]);
						if (n < rows[i].Count - 1)
						{
							sb.Append(" !! ");
						}
					}
				}
				else
				{
					sb.Append(Environment.NewLine);
					sb.Append("|-");
					sb.Append(Environment.NewLine);
					sb.Append("| ");
					for (int n = 0; n < rows[i].Count; n++)
					{
						sb.Append(rows[i][n]);
						if (n < rows[i].Count - 1)
						{
							sb.Append(" || ");
						}
					}
				}
			}

			sb.Append(Environment.NewLine);
			sb.Append("|}");
			return sb.ToString();
		}

		public static string RowsToHtmlTable(List<List<string>> rows)
		{
			StringBuilder sb = new StringBuilder();



		    sb.Append("<!doctype html><html lang=\"en-us\"><head><meta charset=\"utf-8\" /></head><body>");
            sb.Append("<style type='text/css'>\ntable {\n\tborder-collapse: collapse;\n}\n\n");
			sb.Append("table, th, td {\nborder: 1px solid black;\n}\n\n</style>\n");
			sb.Append("<table>");

			for (int i = 0; i < rows.Count; i++)
			{
				if (i == 0)
				{
					sb.Append("\n\t<tr>");

					for (int n = 0; n < rows[i].Count; n++)
					{
						sb.Append("\n\t\t<th>");
						sb.Append("\n\t\t\t" + HttpUtility.HtmlEncode(rows[i][n]));
						sb.Append("\n\t\t</th>");
					}

					sb.Append("\n\t</tr>");
				}
				else
				{
					sb.Append("\n\t<tr>");

					for (int n = 0; n < rows[i].Count; n++)
					{
						sb.Append("\n\t\t<td>");
						sb.Append("\n\t\t\t" + HttpUtility.HtmlEncode(rows[i][n]));
						sb.Append("\n\t\t</td>");
					}

					sb.Append("\n\t</tr>");
				}
			}
            
			sb.Append("\n</table></body></html>");
			return sb.ToString();
		}

		public static string RowsToJSON_ArrayOfArrays(List<List<string>> rows)
		{
			StringBuilder sb = new StringBuilder();
			sb.Append("[\n");

			for (int i = 0; i < rows.Count; i++)
			{
				sb.Append("\n\t[");

				for (int n = 0; n < rows[i].Count; n++)
				{
					if (n == rows[i].Count - 1)
					{
						sb.Append($"\n\t\t\t\"{rows[i][n].Replace("\"", "\\\"")}\"");
					}
					else
					{
						sb.Append($"\n\t\t\t\"{rows[i][n].Replace("\"", "\\\"")}\",");
					}
				}

				sb.Append("\n\t]");

				if (i < rows.Count - 1)
					sb.Append(",");
			}

			sb.Append("\n]");
			return sb.ToString();
		}

		public static string RowsToJSON_ArrayOfObjects(List<List<string>> rows)
		{
			var sb = new StringBuilder();
			sb.Append("[\n");

			var keys = new List<string>();
			for (int n = 0; n < rows[0].Count; n++)
			{
				keys.Add($"{rows[0][n]}");
			}

			for (int i = 1; i < rows.Count; i++)
			{
				sb.Append("\n\t{");

				for (int n = 0; n < rows[i].Count; n++)
				{
					sb.Append($"\n\t\t\"{keys[n].Replace("\"", string.Empty)}\" : \"{rows[i][n].Replace("\"", "\\\"")}\"");
				}

				sb.Append("\n\t}");

				if (i < rows.Count - 1)
					sb.Append(",");
			}

			sb.Append("\n]");
			return sb.ToString();
		}

		public static void RowsToExcel(List<List<string>> rows, string outExcelFileName)
		{
			ExcelAutomate.RowsToExcelFile(rows, outExcelFileName);
		}

		public static void GenerateOutputFile(string format, List<List<string>> rows, string outfile = null)
		{
			string text;
			outfile = outfile ?? GetDefaultOutputFileName(format);
			if(!Path.IsPathRooted(outfile))
			{
				outfile = $"{Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)}\\{outfile}";
			}

			switch (format)
			{
				case "wikitable":
					{
						text = RowsToWikiTable(rows);
						File.WriteAllText(outfile, text);
						break;
					}
				case "jsonobjects":
					{
						text = RowsToJSON_ArrayOfObjects(rows);
						File.WriteAllText(outfile, text);
						break;
					}
				case "jsonarrays":
					{
						text = RowsToJSON_ArrayOfArrays(rows);
						File.WriteAllText(outfile, text);
						break;
					}
				case "excel":
					{
						RowsToExcel(rows, outfile);
						break;
					}
				default:
					{
                        //includes case "html":
                        text = RowsToHtmlTable(rows);
						File.WriteAllText(outfile, text);
						break;
					}
			}
		}

		public static string GetDefaultOutputFileName(string format)
		{
			string s = $"output-{DateTime.Now.ToString("yyyy-MM-dd_HHmmss")}";

			switch(format)
			{
				case "wikitable":
					{
						return $"{s}.txt";
					}
				case "jsonobjects":
					{
						return $"{s}.objects.json";
					}
				case "jsonarrays":
					{
						return $"{s}.arrays.json";
					}
				case "excel":
					{
						return $"{s}.xlsx";
					}
				default:
					{
                        //includes case "html":
                        return $"{s}.html";
					}
			}
		}
 
		public static SimpleArgParser GetArguments(List<SimpleArg> supportedArgs, string[] args)
		{
            //Validate against general rules as specified above.
		    SimpleArgParser parser = new SimpleArgParser(supportedArgs, args);

			//Fix for filename ->abs filename - for case where app is called via batch file, current path is different
			if (!Path.IsPathRooted(parser.ParsedArguments["-filename"]))
			{
				parser.ParsedArguments["-filename"] = Path.GetFullPath(parser.ParsedArguments["-filename"]);
			}

			return parser;
		}
	}
}
