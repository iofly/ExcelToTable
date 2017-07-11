using SimpleArgs;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToTable
{
	public class Utils
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

		public static string RowsToHTMLTable(List<List<string>> rows)
		{
			StringBuilder sb = new StringBuilder();
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
						sb.Append("\n\t\t\t" + rows[i][n]);
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
						sb.Append("\n\t\t\t" + rows[i][n]);
						sb.Append("\n\t\t</td>");
					}

					sb.Append("\n\t</tr>");
				}
			}

			sb.Append("\n</table>");
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
						sb.Append(String.Format("\n\t\t\t\"{0}\"", rows[i][n].Replace("\"", "\\\"")));
					}
					else
					{
						sb.Append(String.Format("\n\t\t\t\"{0}\",", rows[i][n].Replace("\"", "\\\"")));
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
			StringBuilder sb = new StringBuilder();
			sb.Append("[\n");

			List<string> keys = new List<string>();
			for (int n = 0; n < rows[0].Count; n++)
			{
				keys.Add(String.Format("{0}", rows[0][n]));
			}

			for (int i = 1; i < rows.Count; i++)
			{
				sb.Append("\n\t{");

				for (int n = 0; n < rows[i].Count; n++)
				{
					sb.Append(String.Format("\n\t\t\"{0}\" : \"{1}\"", keys[n].Replace("\"", string.Empty), rows[i][n].Replace("\"", "\\\"")));
				}

				sb.Append("\n\t}");

				if (i < rows.Count - 1)
					sb.Append(",");
			}

			sb.Append("\n]");
			return sb.ToString();
		}

		public static void RowsToExcel(List<List<string>> rows, string OutExcelFileName)
		{
			ExcelAutomate.RowsToExcelFile(rows, OutExcelFileName);
		}

		public static void GenerateOutputFile(string format, List<List<string>> rows, string outfile = null)
		{
			string text = null;
			outfile = outfile ?? GetDefaultOutputFileName(format);
			if(!System.IO.Path.IsPathRooted(outfile))
			{
				outfile = String.Format("{0}\\{1}", System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), outfile);
			}

			switch (format)
			{
				case "wikitable":
					{
						text = Utils.RowsToWikiTable(rows);
						System.IO.File.WriteAllText(outfile, text);
						break;
					}
				case "jsonobjects":
					{
						text = Utils.RowsToJSON_ArrayOfObjects(rows);
						System.IO.File.WriteAllText(outfile, text);
						break;
					}
				case "jsonarrays":
					{
						text = Utils.RowsToJSON_ArrayOfArrays(rows);
						System.IO.File.WriteAllText(outfile, text);
						break;
					}
				case "excel":
					{
						Utils.RowsToExcel(rows, outfile);
						break;
					}
				case "html":
				default:
					{
						text = Utils.RowsToHTMLTable(rows);
						System.IO.File.WriteAllText(outfile, text);
						break;
					}
			}
		}

		public static string GetDefaultOutputFileName(string format)
		{
			string s = String.Format("output-{0}", DateTime.Now.ToString("yyyy-MM-dd_HHmmss"));

			switch(format)
			{
				case "wikitable":
					{
						return String.Format("{0}.txt", s);
					}
				case "jsonobjects":
					{
						return String.Format("{0}.objects.json", s);
					}
				case "jsonarrays":
					{
						return String.Format("{0}.arrays.json", s);
					}
				case "excel":
					{
						return String.Format("{0}.xlsx", s);
					}
				case "html":
				default:
					{
						return String.Format("{0}.html", s);
					}
			}
		}
 
		public static SimpleArgParser GetArguments(List<SimpleArg> supportedArgs, string[] args)
		{
			SimpleArgParser parser = null;

			//Validate against general rules as specified above.
			try
			{
				parser = new SimpleArgParser(supportedArgs, args);
			}
			catch (ArgumentException ex)
			{
				throw ex;
			}

			//Fix for filename ->abs filename - for case where app is called via batch file, current path is different
			if (!System.IO.Path.IsPathRooted(parser.ParsedArguments["-filename"]))
			{
				parser.ParsedArguments["-filename"] = System.IO.Path.GetFullPath(parser.ParsedArguments["-filename"]);
			}

			return parser;

		}
	}
}
