using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReaderWPF.CS_Files
{
	public class SheetComparer : IDisposable
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

		readonly private Excel.Application _app;
		readonly private Excel.Workbook _book1;
		readonly private Excel.Workbook _book2;
		readonly private Excel.Workbook _bookOut;
		readonly private Excel.Worksheet _writeSheet;
		readonly private Excel.Range _writeRange;

		public SheetComparer(string file1Path, string file2Path, string fileOutPath)
		{
			_app = new Excel.Application();
			_book1 = _app.Workbooks.Open(file1Path);
			_book2 = _app.Workbooks.Open(file2Path);
			_bookOut = _app.Workbooks.Open(fileOutPath);
			_writeSheet = _bookOut.Sheets[1];
			_writeRange = _writeSheet.UsedRange;
		}
		public void Dispose()
		{
			_book1.Close();
			_book2.Close();
			_bookOut.Save();
			_bookOut.Close();
			_app.Quit();
		}

		public void CompareSheet(ProgressBar progressBar)
		{
			int writeBookRow = 1;

			_writeSheet.Name = _book1.Name;
			_writeRange.Cells[1, 1].Value = "Hall";
			_writeRange.Cells[1, 2].Value = "Jack Type";
			_writeRange.Cells[1, 3].Value = "Jack";

			//for (int i = 0; i < _book1.Worksheets.Count; ++i)
			int i = 0;
			foreach(Excel.Worksheet sheet in _book1.Worksheets)
			{
				progressBar.SheetName = sheet.Name;
				writeBookRow = Compare(sheet, writeBookRow, progressBar);
				i++;
				progressBar.ProgressSheets = 100 * (double)i / (_book1.Worksheets.Count - 1);
			}
		}

		public int Compare(Excel.Worksheet sheet, int writeBookRow, ProgressBar progressBar)
		{
			//update current sheet name in text box here
			int writebookColumns = 1;
			bool written;
			Excel.Worksheet compareSheet = _book2.Sheets[sheet.Index];
			Excel.Range compareRange = compareSheet.UsedRange;
			Excel.Range range = sheet.UsedRange;
			int rows = range.Rows.Count;
			int columns = range.Columns.Count;
			for (int i = 1; i <= rows; i++)
			{
				written = false;
				float percentage = ((float)i / rows) * 100;
				progressBar.Progress = percentage;
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
							_writeRange.Cells[writeBookRow, 1].Value = sheet.Name;
							_writeRange.Cells[writeBookRow, 2].Value = range.Cells[1, j].Value;
							_writeRange.Cells[writeBookRow, 3].Value = range.Cells[i, 4].Value;
							_writeRange.Cells[writeBookRow, writebookColumns + 3].Value = range.Cells[i, j].Value;
							writebookColumns++;
							written = true;
						}
					}
					if (written)
						writeBookRow++;
					writebookColumns = 1;
				}
			}
			return writeBookRow;
		}
	}
}
