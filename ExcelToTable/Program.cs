﻿using System;
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

            List<SimpleArg> supportedArgs = new List<SimpleArg>();
            supportedArgs.Add(new SimpleArg { Name = "-filename", IsSwitch = false, Required = true, DefaultValue = null, ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-format", IsSwitch = false, Required = true, DefaultValue = "html", ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-range", IsSwitch = false, Required = true, DefaultValue = null, ArgType = SimpleArgType.String });
            supportedArgs.Add(new SimpleArg { Name = "-worksheet", IsSwitch = false, Required = false, DefaultValue = 1, ArgType = SimpleArgType.Integer });
            supportedArgs.Add(new SimpleArg { Name = "-fromdate", IsSwitch = false, Required = false, DefaultValue = DateTime.MinValue, ArgType = SimpleArgType.DateTime });
            supportedArgs.Add(new SimpleArg { Name = "-v", IsSwitch = true, Required = false });
            supportedArgs.Add(new SimpleArg { Name = "-speed", IsSwitch = true, Required = false });
            supportedArgs.Add(new SimpleArg { Name = "-hello", IsSwitch = false, Required = false, DefaultValue = null, ArgType = SimpleArgType.String });

            Dictionary<string, dynamic> ar = new Dictionary<string, dynamic>();
            try
            {
                ar = SimpleArgParser.ParseArgs(args, supportedArgs);
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Error: {0}", ex.Message));
                ShowUsage();
                return;
            }

            Console.WriteLine("Successfully parsed args");
            if (!System.IO.Path.IsPathRooted(ar["-filename"]))
            {
                ar["-filename"] = System.IO.Path.GetFullPath(ar["-filename"]);
            }

         
            if (!ValidateArgs(ar))
            {
                return;
            }


            rows = ReadExcelRows(ar["-filename"], out rc, out ResultDesc, ar["-worksheet"]);
            if (rc == ResultCode.ErrorOpeningFile)
            {
                Console.WriteLine(String.Format("ErrorOpeningFile: {0}", ResultDesc));
                return;
            }
            


            if (ar["-format"] == "wikitable")
            {
                text = RowsToWikiTable(rows);
                fileext = ".txt";
            }
            else
            {
                text = RowsToHTMLTable(rows);
                fileext = ".html";
            }

            System.IO.File.WriteAllText(System.IO.Path.GetFileName(ar["-filename"]) + fileext, text);
        }

        #region Argument stuff
        private static void ShowUsage()
        {
            Console.WriteLine(String.Format("Usage: {0} [excelfilename] worksheet=[1-n] format=[html|wikitable]", System.AppDomain.CurrentDomain.FriendlyName));
            Console.WriteLine("worksheet is optional. Defaults to 1");
            Console.WriteLine("format is optional. Defaults to wikitable");
            Console.WriteLine(Environment.NewLine);
        }

        private static Dictionary<string, string> ParseArgs(string[] args)
        {
            Dictionary<string, string> parsedArgs = new Dictionary<string, string>();
            //preload optional args
            parsedArgs.Add("worksheet", "1");
            parsedArgs.Add("format", "wikitable");

            foreach (string s in args)
            {
                var match = System.Text.RegularExpressions.Regex.Match(s, "worksheet=([0-9])");
                if (match.Success)
                {
                    parsedArgs["worksheet"] = match.Groups[1].Value;
                    continue;
                }
                else if (System.IO.File.Exists(s))
                {
                    parsedArgs.Add("inputfile", s);
                    continue;
                }

                match = System.Text.RegularExpressions.Regex.Match(s, "format=(html|wikitable)");
                if (match.Success)
                {
                    parsedArgs["format"] = match.Groups[1].Value;
                    continue;
                }
            }

            return parsedArgs;
        }

        private static bool ValidateArgs(Dictionary<string, dynamic> ar)
        {
            if (!ar.ContainsKey("-filename"))
            {
                Console.WriteLine("Missing parameter -filename : Must provide input file path");
                ShowUsage();
                return false;
            }
            else if (!System.IO.File.Exists(ar["-filename"]))
            {
                Console.WriteLine(String.Format("Input file '{0}' does not exist", ar["-filename"]));
                ShowUsage();
                return false;
            }

            if ("html|wikitable".IndexOf(ar["-format"]) < 0)
            {
                Console.WriteLine(String.Format("Output format '{0}' is not valid", ar["-format"]));
                return false;
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

        #endregion

    }
}
