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
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReaderWPF
{
	/// <summary>
	/// Interaction logic for ProgressBar.xaml
	/// </summary>
	public partial class ProgressBar : Page
	{
		string File1Path;
		string File2Path;
		string FileOutPath;
		int FileOutIndex;
		public ProgressBar()
		{
			InitializeComponent();
		}
		public ProgressBar(string fp1, string fp2, string fpo, string foi):this()
		{
			File1Path = fp1;
			File2Path = fp2;
			FileOutPath = fpo;
			FileOutIndex = Convert.ToInt32(foi);
			//this.Loaded += new RoutedEventHandler(OnLoad);
		}
	}
}
