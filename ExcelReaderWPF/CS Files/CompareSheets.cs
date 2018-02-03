using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReaderWPF.CS_Files
{
	class CompareSheets
	{
		static readonly List<string> dictionary = new List<string>()
		{
			"CABLE BOX",
			"FACE-PLATE",
			"WIRE-MOLD",
			"CABLE TV",
			"DATA JACK",
			"PHONE JACK",
			"RED PHONE",
			"CABLE " + "\n" + "BOX ", //Replace these two with a regex attempt to find all cable box?
			"CABLE" + "\n" + "BOX ",
		};

		public static void openExcel(string file1Path, string file2Path, string fileOutPath)
		{
			int writeBookRow = 1;
			Excel.Application app = new Excel.Application();
			Excel.Workbook book1 = app.Workbooks.Open(file1Path);
			Excel.Workbook book2 = app.Workbooks.Open(file2Path);
			Excel.Workbook bookOut = app.Workbooks.Open(fileOutPath);

			Excel.Worksheet writeSheet = bookOut.Sheets[1];
			Excel.Range writeRange = writeSheet.UsedRange;

			writeSheet.Name = book1.Name;
			writeRange.Cells[1, 1].Value = "Hall";
			writeRange.Cells[1, 2].Value = "Jack Type";
			writeRange.Cells[1, 3].Value = "Jack";

			foreach(Excel.Worksheet sheet in book1.Worksheets)
			{
				writeBookRow = Compare(sheet, book2, writeRange, writeBookRow);
			}
			book1.Close();
			book2.Close();
			bookOut.Save();
			bookOut.Close();
			app.Quit();
		}

		//public static int Compare(Excel.Worksheet sheet, Excel.Workbook workbook2, Excel.Range writeRange, int writebookRow)
		public static int Compare(Excel.Worksheet sheet, Excel.Workbook workbook2, Excel.Range writeRange, int writebookRow)
		{
			int writebookColumns = 1;
			bool written;
			Excel.Worksheet compareSheet = workbook2.Sheets[sheet.Index];
			Excel.Range compareRange = compareSheet.UsedRange;
			Excel.Range range = sheet.UsedRange;
			int rows = range.Rows.Count;
			int columns = range.Columns.Count;
			for (int i = 1; i <= rows; i++)
			{
				written = false;
				float percentage = ((float)i / rows) * 100;
				//Console.Write(sheet.Name + ": " + percentage.ToString("0.0") + "%");
				if (range.Cells[i, 4].Value == compareRange.Cells[i, 4].Value && range.Cells[i, 4].Value != null)
				{
					for (int j = 5; j <= 11; j++)
					{
						if (range.Cells[i, j].Value == compareRange.Cells[i, j].Value && range.Cells[i, j].Value != null)
						{
							if (dictionary.Contains(Convert.ToString(range.Cells[i, j].Value)) || int.TryParse(Convert.ToString(range.Cells[i, j].Value), out int n))
							{
								break;
							}
							writeRange.Cells[writebookRow, 1].Value = sheet.Name;
							writeRange.Cells[writebookRow, 2].Value = range.Cells[1, j].Value;
							writeRange.Cells[writebookRow, 3].Value = range.Cells[i, 4].Value;
							writeRange.Cells[writebookRow, writebookColumns + 3].Value = range.Cells[i, j].Value;
							writebookColumns++;
							written = true;
						}
					}
					if (written)
						writebookRow++;
					writebookColumns = 1;
				}
			}
			return writebookRow;
		}
	}
}
