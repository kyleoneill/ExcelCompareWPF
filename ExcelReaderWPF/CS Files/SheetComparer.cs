using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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

		//Compares user input index to the current amount of sheets in the output file. If the index is higher than the amount of sheets, creates a new sheet and tests again.
		public void AddIndex(Excel.Workbook book, int index)
		{
			if(index > book.Sheets.Count)
			{
				book.Sheets.Add(After: book.Sheets[book.Sheets.Count]);
				AddIndex(book, index);
			}
		}

		//Class constructor. Opens two workbook objects for the files to be read and one excel range project for the sheet on the book that will be written to.
		public SheetComparer(string file1Path, string file2Path, string fileOutPath, int index)
		{
			_app = new Excel.Application();
			_book1 = _app.Workbooks.Open(file1Path);
			_book2 = _app.Workbooks.Open(file2Path);
			_bookOut = _app.Workbooks.Open(fileOutPath);
			AddIndex(_bookOut, index);
			_writeSheet = _bookOut.Sheets[index];
			_writeRange = _writeSheet.UsedRange;
		}

		//Dispose function closes excel books and excel application when the task using this class ends or when called.
		public void Dispose()
		{
			_book1.Close(false); //False prevents excel from asking if the user wants to save changes
			_book2.Close(false);
			_bookOut.Save();
			_bookOut.Close();
			_app.Quit();
		}

		//main function of the class. Compares all sheets in two workbooks and outputs data onto the output sheet chosen by the user.
		//Output data includes any data from matching jacks that is the same between the two input data books.
		public void CompareSheet(ProgressBar progressBar)
		{
			//TryCatch ends program if the user attemps to enter the same input multiple times.
			try
			{
				_writeSheet.Name = _book1.Name;
			}
			catch
			{
				MessageBoxResult valueTooHigh = MessageBox.Show("The selected quad already exists in the output file.", "Error", MessageBoxButton.OK);
				Dispose();
				Environment.Exit(0);
			}
			_writeRange.Cells[1, 1].Value = "Hall";
			_writeRange.Cells[1, 2].Value = "Jack Type";
			_writeRange.Cells[1, 3].Value = "Jack";
			_writeRange.Cells[1, 4].Value = "Issue";
			int writeBookRow = 2;

			int i = 0;
			foreach(Excel.Worksheet sheet in _book1.Worksheets)
			{
				progressBar.SheetName = sheet.Name;
				writeBookRow = Compare(sheet, writeBookRow, progressBar);
				i++;
				progressBar.ProgressSheets = 100 * (double)i / (_book1.Worksheets.Count - 1);
			}
		}

		//Compares the names of sheets in both of the input books to make sure that the correct sheets are paired up, even if the sheets are out of order in the books
		public Excel.Worksheet GetSheet(Excel.Worksheet sheet)
		{
			foreach(Excel.Worksheet sheet2 in _book2.Sheets)
			{
				if(sheet.Name == sheet2.Name)
				{
					return sheet2;
				}
			}
			MessageBoxResult valueTooHigh = MessageBox.Show("Sheet '" + sheet.Name + "' does not exist in book " + _book2.Name, "Error", MessageBoxButton.OK);
			Dispose();
			Environment.Exit(0);
			return null;
		}

		//comparison for each sheet in a book
		public int Compare(Excel.Worksheet sheet, int writeBookRow, ProgressBar progressBar)
		{
			try
			{
				int writebookColumns = 1;
				bool written;
				Excel.Worksheet compareSheet = GetSheet(sheet);
				Excel.Range compareRange = compareSheet.UsedRange;
				Excel.Range range = sheet.UsedRange;
				int rows = range.Rows.Count;
				int columns = range.Columns.Count;
				for (int i = 1; i <= rows; i++)
				{
					written = false; //bool value is used to increment the current column on the output sheet when data is written
					float percentage = ((float)i / rows) * 100;
					progressBar.Progress = percentage;
					//If jack number on input book 1 == jack number on input book 2 and the value is not null
					if (Convert.ToString(range.Cells[i, 4].Value) == Convert.ToString(compareRange.Cells[i, 4].Value) && range.Cells[i, 4].Value != null)
					{
						//For each column of jack information
						for (int j = 5; j <= 11; j++)
						{
							//If the columns are the same between input book 1 and input book 2 and the columns are not null
							if (range.Cells[i, j].Value == compareRange.Cells[i, j].Value && range.Cells[i, j].Value != null)
							{
								//Checks the above dictionary to get rid of garbage data
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
			catch
			{
				MessageBoxResult valueTooHigh = MessageBox.Show("One/both of the inserted sheets are using an unsupported format.", "Error", MessageBoxButton.OK);
				Dispose();
				Environment.Exit(0);
				return 0;
			}
		}
	}
}
