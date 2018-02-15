using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelReaderWPF.CS_Files;

namespace ExcelReaderWPF
{
	/// <summary>
	/// Interaction logic for ProgressBar.xaml
	/// </summary>
	
	// Two progress bars - one for current row/all rows in a sheet, one for current sheet(index)/all sheets in a book

	public partial class ProgressBar : Page
	{
		string File1Path;
		string File2Path;
		string FileOutPath;
		int FileOutIndex;

		Task _compareTask;

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
		public ProgressBar(string fp1, string fp2, string fpo, int foi):this()
		{
			File1Path = fp1;
			File2Path = fp2;
			FileOutPath = fpo;
			FileOutIndex = foi;

			_compareTask = Task.Run(() =>
			{
				using (var comparer = new SheetComparer(File1Path, File2Path, FileOutPath, FileOutIndex)) //After the comparer obj is done being used, the dispose function is called
				{
					comparer.CompareSheet(this);
					comparer.Dispose();
				}
			});
		}
	}
}
