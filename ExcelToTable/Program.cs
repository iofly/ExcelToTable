using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;

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

            #region Define permitted arguments
            List<SimpleArg> supportedArgs = new List<SimpleArg>();
            supportedArgs.Add(new SimpleArg { Name = "-filename", IsSwitch = false, Required = true, DefaultValue = null, ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-format", IsSwitch = false, Required = true, DefaultValue = "html", ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-worksheet", IsSwitch = false, Required = false, DefaultValue = 1, ArgType = SimpleArgType.Integer });

            supportedArgs.Add(new SimpleArg { Name = "-range", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-fromdate", IsSwitch = false, Required = false, DefaultValue = DateTime.MinValue, ArgType = SimpleArgType.DateTime });
            supportedArgs.Add(new SimpleArg { Name = "-v", IsSwitch = true, Required = false });
            supportedArgs.Add(new SimpleArg { Name = "-speed", IsSwitch = true, Required = false });
            supportedArgs.Add(new SimpleArg { Name = "-hello", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String });
            #endregion

            #region Parse  and validate arguments
            
            SimpleArgParser parser = new SimpleArgParser(supportedArgs);

            Dictionary<string, dynamic> ar = new Dictionary<string, dynamic>();
            try
            {
                ar = parser.ParseArgs(args);
                parser.ValidateArgs(ar);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                ShowUsage();
                return;
            }

            //Fix for filename ->abs filename - for case where app is called via batch file
            if (!System.IO.Path.IsPathRooted(ar["-filename"]))
            {
                ar["-filename"] = System.IO.Path.GetFullPath(ar["-filename"]);
            }

            try
            {
                ValidateAppArgs(ar);
            }
            catch(ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                ShowUsage();
                return;
            }
            #endregion

            #region Read excel file
            rows = ReadExcelRows(ar["-filename"], out rc, out ResultDesc, ar["-worksheet"]);
            if (rc == ResultCode.ErrorOpeningFile)
            {
                Console.WriteLine(String.Format("ErrorOpeningFile: {0}", ResultDesc));
                return;
            }
            #endregion

            #region Output result
            if (ar["-format"] == "wikitable")
            {
                text = RowsToWikiTable(rows);
                fileext = ".txt";
            }
            else if (ar["-format"] == "jsonsobjects")
            {
                text = RowsToJSON_ArrayOfObjects(rows);
                fileext = ".objects.json";
            }
            else if (ar["-format"] == "jsonarrays")
            {
                text = RowsToJSON_ArrayOfArrays(rows);
                fileext = ".arrays.json";
            }
            else
            {
                text = RowsToHTMLTable(rows);
                fileext = ".html";
            }

            

            System.IO.File.WriteAllText(System.IO.Path.GetFileName(ar["-filename"]) + fileext, text);
            #endregion
        }

        #region Argument stuff
        private static void ShowUsage()
        {
            Console.WriteLine(String.Format("Usage: {0} [excelfilename] worksheet=[1-n] format=[html|wikitable|jsonarrays|jsonsobjects]", System.AppDomain.CurrentDomain.FriendlyName));
            Console.WriteLine("worksheet is optional. Defaults to 1");
            Console.WriteLine("format is optional. Defaults to wikitable");
            Console.WriteLine(Environment.NewLine);
        }

        private static bool ValidateAppArgs(Dictionary<string, dynamic> ar)
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

        #region Excel stuff
        public static List<List<string>> ReadExcelRows(string ExcelFileName, out ResultCode ResultCode, out string ResultDesc, int worksheet = 1)
		{
            List<List<string>> tablerows = new List<List<string>>();
            Application xlApp = new Application();
            Workbook xlWorkBook = null;
            Worksheet xlWorkSheet = null;

            try
            {
                xlWorkBook = xlApp.Workbooks.Open(ExcelFileName);
            }
            catch(Exception ex)
            {
                xlApp.Quit();
                ReleaseObject(xlApp);
                ResultCode = ResultCode.ErrorOpeningFile;
                ResultDesc = ex.Message;
                return tablerows;
            }

            try
            {
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(worksheet);
                Range range = xlWorkSheet.UsedRange; //only look at area where there is data
                List<string> cells;
                string pw = string.Empty;
                int c = 1, rCnt = 0;

                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {
                    cells = new List<string>();
                    for (c = 1; c <= range.Columns.Count; c++)
                    {
                        //cells.Add(String.Format("{0}", (range.Cells[rCnt, c] as Range).Value2));
                        cells.Add(String.Format("{0}", GetRangeStr(range.Cells[rCnt, c] as Range)));
                    }

                    tablerows.Add(cells);
                }
            }
            catch(Exception ex1)
            {
                ResultCode = ResultCode.ErrorOpeningFile;
                ResultDesc = ex1.Message;
                return tablerows;
            }
            finally
            {
                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
                ReleaseObject(xlWorkSheet);
                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
            }
            ResultCode = ResultCode.Success;
            ResultDesc = "Success";
            return tablerows;
		}

        private static string GetRangeStr(Range range)
        {
            if (range == null) return string.Empty;
            else
            {
                if(range.Value2 == null)
                {
                    return string.Empty;
                }
                else
                {
                    return String.Format("{0}", range.Value2);
                }
            }
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion

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
                    if(n == rows[i].Count - 1)
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

    }
}
