using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using SimpleArgs;

namespace ExcelToTable
{
	public class ExcelAutomate
	{
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
					//Cell index is [Y,X]! wtf
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
				xlWorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
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

		public static void RowsToExcelFile(List<List<string>> ExcelData, string OutExcelFileName)
		{
			var xlApp = new Microsoft.Office.Interop.Excel.Application();
			Workbook xlWorkBook = null;
			Worksheet xlWorkSheet = null;

			if (xlApp == null)
			{
				ReleaseObject(xlApp);
				throw new Exception("EXCEL could not be started. Check that your office installation and project references are correct.");
			}

			try
			{ 
				xlApp.Visible = false;
				xlWorkBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
				xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				xlWorkSheet.Name = "Exported";

				Range NewCell;
				int i = 3;
				int rowindex = 0;
				int colindex = 0;

				foreach (var row in ExcelData)
				{
					colindex = 0;
					foreach (var col in row)
					{
						NewCell = (Range)xlWorkSheet.Cells[rowindex + 1, colindex + 1];
						NewCell.Value = col;
						NewCell.Font.Bold = rowindex == 0;
						colindex++;
					}
					i++;
					rowindex++;
				}

				Range used_range = xlWorkSheet.UsedRange;
				used_range.Columns.AutoFit();
				xlWorkBook.SaveAs(OutExcelFileName);
			}
			catch
			{
				
			}
			finally
			{
				xlWorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
				xlApp.Quit();
				ReleaseObject(xlWorkSheet);
				ReleaseObject(xlWorkBook);
				ReleaseObject(xlApp);
			}
		}
	}
}
