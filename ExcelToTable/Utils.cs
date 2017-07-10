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

        public static void GenerateOutputFile(string format, List<List<string>> rows, string inputfile, string outfile = null)
        {
            string text = null;
            string fileext = null;
            if (format == "wikitable")
            {
                text = Utils.RowsToWikiTable(rows);
                fileext = ".txt";
            }
            else if (format == "jsonobjects")
            {
                text = Utils.RowsToJSON_ArrayOfObjects(rows);
                fileext = ".objects.json";
            }
            else if (format == "jsonarrays")
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
                System.IO.File.WriteAllText(System.IO.Path.GetFileName(inputfile) + fileext, text);
            }
        }

        /// <summary>
        /// App specific validation of arguments
        /// </summary>
        /// <param name="ar">Passed arguments</param>
        /// <returns></returns>
        public static bool ValidateAppArgs(Dictionary<string, dynamic> ar)
        {
            if(ar.ContainsKey("-format"))
            {
                if ("html|wikitable|jsonarrays|jsonobjects".IndexOf(ar["-format"]) < 0)
                {
                    throw new ArgumentException(String.Format("Output format '{0}' is not valid", ar["-format"]));
                }
            }

  

            return true;
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
                Console.WriteLine(String.Empty);
                Console.WriteLine(String.Format("Error: {0}", ex.Message));
                SimpleArgParser.ShowUsage(supportedArgs);
                throw ex;
            }

            //App specific validation of args
            try
            {
                Utils.ValidateAppArgs(parser.ParsedArguments);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                parser.ShowUsage();
                throw ex;
            }

            //Fix for filename ->abs filename - for case where app is called via batch file
            if (!System.IO.Path.IsPathRooted(parser.ParsedArguments["-filename"]))
            {
                parser.ParsedArguments["-filename"] = System.IO.Path.GetFullPath(parser.ParsedArguments["-filename"]);
            }

            string outfile = string.Empty;
            if (parser.ParsedArguments.ContainsKey("-outfile"))
            {
                outfile = System.IO.Path.GetFullPath(parser.ParsedArguments["-outfile"]);
            }

            WorkSheetRangeCoordinates wsrc = null;
            if (parser.ParsedArguments.ContainsKey("-range"))
            {
                wsrc = SimpleArgParser.ParseExcelRange(parser.ParsedArguments["-range"]);
                if (wsrc == null)
                {
                    Console.WriteLine(String.Empty);
                    Console.WriteLine(String.Format("Error: Range parameter is not valid: -range => {0}", parser.ParsedArguments["-range"]));
                    SimpleArgParser.ShowUsage(supportedArgs);
                    return null;
                }
            }

            return parser;

        }
    }
}
