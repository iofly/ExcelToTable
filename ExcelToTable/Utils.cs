using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Threading.Tasks;

namespace ExcelToTable
{
    public class Utils
    {
        
        #region Output calls
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

        #endregion

        #region Argument stuff
        public static void ShowUsage()
        {
            Console.WriteLine(String.Format("Usage: {0} -filename [excelfilename] -worksheet [1-n] -format [html|wikitable|jsonarrays|jsonsobjects]", System.AppDomain.CurrentDomain.FriendlyName));
            Console.WriteLine("-filename: The Excel file name, relative or absolute path.");
            Console.WriteLine("-worksheet is optional. Defaults to 1");
            Console.WriteLine("-format [optional]. Output file format. Defaults to wikitable");
            Console.WriteLine("\t\tjsonarrays option outputs each row as an array, making a 2-d array.");
            Console.WriteLine("\t\tjsonsobjects option outputs an array of objects using the first row as object key names");
            Console.WriteLine("-range [optional]. Excel worksheet range to export. Defaults to used range");
            Console.WriteLine(Environment.NewLine);
        }

        /// <summary>
        /// App specific validation of arguments
        /// </summary>
        /// <param name="ar">Passed arguments</param>
        /// <returns></returns>
        public static bool ValidateAppArgs(Dictionary<string, dynamic> ar)
        {
            if (!ar.ContainsKey("-filename"))
            {
                throw new ArgumentException("Missing parameter -filename : Must provide input file path");
            }
            else if (!System.IO.File.Exists(ar["-filename"]))
            {
                throw new ArgumentException(String.Format("Input file '{0}' does not exist", ar["-filename"]));
            }
            else if ("html|wikitable|jsonarrays|jsonsobjects".IndexOf(ar["-format"]) < 0)
            {
                throw new ArgumentException(String.Format("Output format '{0}' is not valid", ar["-format"]));
            }

            return true;
        }

        #endregion

    
    }

    public class WorkSheetCoordinate
    {
        public int X { get; set; }
        public int Y { get; set; }
    }

    public class WorkSheetRangeCoordinates
    {
        public WorkSheetRangeCoordinates()
        {
            _TopLeft = new WorkSheetCoordinate();
            _BottomRight = new WorkSheetCoordinate();

        }

        private WorkSheetCoordinate _TopLeft;
        public WorkSheetCoordinate TopLeft
        {
            get
            {
                return _TopLeft;
            }
            set
            {
                _TopLeft = value;
            }
        }
        private WorkSheetCoordinate _BottomRight;
        public WorkSheetCoordinate BottomRight
        {
            get
            {
                return _BottomRight;
            }
            set
            {
                _BottomRight = value;
            }
        }
    }
}
