using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ExcelToTable
{
	public class ExcelAutomate
	{
		public static List<List<string>> ReadExcelRows(string excelFileName, out ResultCode resultCode, 
			out string resultDesc, int worksheet = 1, WorkSheetRangeCoordinates wsrc = null)
		{
			List<List<string>> tablerows = new List<List<string>>();
			Application xlApp = new Application();
			Workbook xlWorkBook;
			Worksheet xlWorkSheet = null;

			try
			{
				xlWorkBook = xlApp.Workbooks.Open(excelFileName);
			}
			catch (Exception ex)
			{
				xlApp.Quit();
				ReleaseObject(xlApp);
				resultCode = ResultCode.ErrorOpeningFile;
				resultDesc = ex.Message;
				return tablerows;
			}

			try
			{
				xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(worksheet);
				Range range;
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
					range = xlWorkSheet.Range[c1, c2];
				}

				int  rowIndex;

				for (rowIndex = 1; rowIndex <= range.Rows.Count; rowIndex++)
				{
					List<string> cells = new List<string>();
					for (int colIndex = 1; colIndex <= range.Columns.Count; colIndex++)
					{
						//cells.Add(String.Format("{0}", (range.Cells[rCnt, c] as Range).Value2));
						cells.Add($"{GetRangeStr(range.Cells[rowIndex, colIndex] as Range)}");
					}

					tablerows.Add(cells);
				}
			}
			catch (Exception ex1)
			{
				resultCode = ResultCode.ErrorOpeningFile;
				resultDesc = ex1.Message;
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
			resultCode = ResultCode.Success;
			resultDesc = "Success";
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
				//obj = null;
			}
			catch (Exception ex)
			{
				//obj = null;
				Console.WriteLine("Unable to release the Object " + ex.Message);
			}
			finally
			{
				GC.Collect();
			}
		}

		public static void RowsToExcelFile(List<List<string>> excelData, string outExcelFileName)
		{
			var xlApp = new Application();
			Workbook xlWorkBook = null;
			Worksheet xlWorkSheet = null;

			try
			{ 
				xlApp.Visible = false;
				xlWorkBook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
				xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				xlWorkSheet.Name = "Exported";

				int rowindex = 0;

				foreach (var row in excelData)
				{
					int colindex = 0;
					foreach (var col in row)
					{
						var newCell = (Range)xlWorkSheet.Cells[rowindex + 1, colindex + 1];
						newCell.Value = col;
						newCell.Font.Bold = rowindex == 0;
						colindex++;
					}
					rowindex++;
				}

				Range usedRange = xlWorkSheet.UsedRange;
				usedRange.Columns.AutoFit();
				xlWorkBook.SaveAs(outExcelFileName);
			}
			finally
			{
				xlWorkBook?.Close(Type.Missing, Type.Missing, Type.Missing);
				xlApp.Quit();
				ReleaseObject(xlWorkSheet);
				ReleaseObject(xlWorkBook);
				ReleaseObject(xlApp);
			}
		}
	}
}
