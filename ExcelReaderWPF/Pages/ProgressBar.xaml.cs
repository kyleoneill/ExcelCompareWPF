using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using ExcelReaderWPF.CS_Files;
using ExcelReaderWPF.Pages;

namespace ExcelReaderWPF
{
	/// <summary>
	/// Interaction logic for ProgressBar.xaml
	/// </summary>
	
	// Two progress bars - one for current row/all rows in a sheet, one for current sheet(index)/all sheets in a book

	public partial class ProgressBar : Page
	{
		public double Progress
		{
			get { return ProgressBarObj.Value; }
			set { Dispatcher.Invoke(() => ProgressBarObj.Value = value); } //Dispatcher sets the value of the progress bar on the same thread that owns the progress bar
		}

		public double ProgressSheets
		{
			get { return ProgressBarSheets.Value; }
			set { Dispatcher.Invoke(() => ProgressBarSheets.Value = value); }
		}

		public string SheetName
		{
			get { return ProgressTextBox.Text; }
			set { Dispatcher.Invoke(() => ProgressTextBox.Text = value); }
		}

		public ProgressBar()
		{
			InitializeComponent();
		}

		public void Compare(string file1Path, string file2Path, string fileOutPath, int fileOutIndex)
		{
			using (var comparer = new SheetComparer(file1Path, file2Path, fileOutPath, fileOutIndex)) //After the comparer obj is done being used, the dispose function is called
			{
				comparer.CompareSheet(this);
			}
		}

		public ProgressBar(string fp1, string fp2, string fpo, int foi):this()
		{
            Task.Run(() =>
            {
                Compare(fp1, fp2, fpo, foi);
                // Navigate back to the first page
                Dispatcher.Invoke(() => NavigationService.Navigate(new FirstPage()));
            });
		}
		public ProgressBar(string folderPath1, string folderPath2, string filePathOut):this()
		{
            Task.Run(() =>
            {
                var firstAssess = Directory.GetFiles(folderPath1, "*.xlsx")
                    .Select(Path.GetFileNameWithoutExtension)
                    .ToArray();
                var secondAssess = Directory.GetFiles(folderPath2, "*.xlsx")
                    .Select(Path.GetFileNameWithoutExtension)
                    .ToArray();

                var fileIndex = 1;
                foreach (var file in firstAssess)
                {
                    var file1Path = folderPath1 + "\\" + file + ".xlsx";
                    var file2Path = folderPath2 + "\\" + (secondAssess[Array.IndexOf(secondAssess, file + " Second")]) + ".xlsx";

                    Compare(file1Path, file2Path, filePathOut, fileIndex);

                    ++fileIndex;
                }

                // Navigate back to the first page
                Dispatcher.Invoke(() => NavigationService.Navigate(new FirstPage()));
            });
		}
	}
}
