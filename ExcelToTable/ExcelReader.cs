using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ExcelToTable
{
    public class ExcelReader
    {
        const string Letters = "_ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static List<List<string>> ReadExcelRows(string ExcelFileName, out ResultCode ResultCode, out string ResultDesc, int worksheet = 1, WorkSheetRangeCoordinates wsrc = null)
        {
            List<List<string>> tablerows = new List<List<string>>();
            Application xlApp = new Application();
            Workbook xlWorkBook = null;
            Worksheet xlWorkSheet = null;

            try
            {
                xlWorkBook = xlApp.Workbooks.Open(ExcelFileName);
            }
            catch (Exception ex)
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
                Range range = null;
                if (wsrc==null)
                {
                    range = xlWorkSheet.UsedRange; //only look at area where there is data
                }
                else
                {
                    Range c1 = xlWorkSheet.Cells[wsrc.TopLeft.Y, wsrc.TopLeft.X];
                    Range c2 = xlWorkSheet.Cells[wsrc.BottomRight.Y, wsrc.BottomRight.X];
                    //oRange = (Excel.Range)oSheet.get_Range(c1, c2);
                    range = xlWorkSheet.get_Range(c1, c2);
                }

                List<string> cells;
                string pw = string.Empty;
                int colIndex = 1, rowIndex = 0;

                for (rowIndex = 1; rowIndex <= range.Rows.Count; rowIndex++)
                {
                    cells = new List<string>();
                    for (colIndex = 1; colIndex <= range.Columns.Count; colIndex++)
                    {
                        //cells.Add(String.Format("{0}", (range.Cells[rCnt, c] as Range).Value2));
                        cells.Add(String.Format("{0}", GetRangeStr(range.Cells[rowIndex, colIndex] as Range)));
                    }

                    tablerows.Add(cells);
                }
            }
            catch (Exception ex1)
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
            if (range == null)
                return string.Empty;
            else
            {
                if (range.Value2 == null)
                {
                    //return range.Value;
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


        public static WorkSheetRangeCoordinates ParseExcelRange(string Range)
        {
            //BMZ4:BNC14
            string regexPattern = @"([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})";
            WorkSheetRangeCoordinates rangeCoords = new WorkSheetRangeCoordinates();
            var matches = Regex.Matches(Range, regexPattern);

            int colIndex = 0;
            int letterIndex = 0;

            if (matches.Count > 0)
            {

                string part1 = matches[0].Groups[1].Value;
                char[] part1X = part1.ToCharArray();
                Array.Reverse(part1X); //make first car the LSC
                for (int i = 0; i < part1X.Length; i++)
                {
                    letterIndex = Letters.IndexOf(part1X[i]);
                    colIndex += letterIndex *  Convert.ToInt32(Math.Pow(26, i));
                }
                rangeCoords.TopLeft.X = colIndex;
                rangeCoords.TopLeft.Y = int.Parse(matches[0].Groups[2].Value);

                colIndex = 0;

                string part2 = matches[0].Groups[3].Value;
                char[] part2X = part2.ToCharArray();
                Array.Reverse(part2X); //make first car the LSC
                for (int i = 0; i < part2X.Length; i++)
                {
                    letterIndex = Letters.IndexOf(part2X[i]);
                    colIndex += letterIndex * Convert.ToInt32(Math.Pow(26, i));
                }
                rangeCoords.BottomRight.X = colIndex;
                rangeCoords.BottomRight.Y = int.Parse(matches[0].Groups[4].Value);
                return rangeCoords;
            }
            else
            {
                return null;
            }
        }

    }
}
