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
using System.Text.RegularExpressions;

namespace ExcelReaderWPF
{
	/// <summary>
	/// Interaction logic for MainPage.xaml
	/// </summary>
	public partial class MainPage : Page
	{
		public MainPage()
		{
			InitializeComponent();
		}
		private static void FileSelect(TextBox textbox)
		{
			var fileDialog = new System.Windows.Forms.OpenFileDialog();
			var result = fileDialog.ShowDialog();
			switch (result)
			{
				case System.Windows.Forms.DialogResult.OK:
					var file = fileDialog.FileName;
					textbox.Text = file;
					textbox.ToolTip = file;
					break;
				case System.Windows.Forms.DialogResult.Cancel:
				default:
					textbox.Text = null;
					textbox.ToolTip = null;
					break;
			}
		}
		private void File1Select_Click(object sender, RoutedEventArgs e)
		{
			FileSelect(File1Box);
		}
		private void File2Select_Click(object sender, RoutedEventArgs e)
		{
			FileSelect(File2Box);
		}

		private void File3Select_Click(object sender, RoutedEventArgs e)
		{
			FileSelect(FileOutBox);
		}
		private static bool IsTextAllowed(string text)
		{
			Regex regex = new Regex("[^0-9.-]+"); //matches disallowed text
			return !regex.IsMatch(text);
		}
		private void FileIndexBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
		{
			e.Handled = !IsTextAllowed(e.Text);
		}
		public void SubmitDataButton_Click(object sender, RoutedEventArgs e)
		{
			if (File1Box.Text != "" && File2Box.Text != "" && FileOutBox.Text != "" && FileIndexBox.Text != "")
			{
				ProgressBar p = new ProgressBar(File1Box.Text, File2Box.Text, FileOutBox.Text, FileIndexBox.Text);
				this.NavigationService.Navigate(p);
			}
		}
	}
}
