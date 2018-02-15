using System;
using System.IO;
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

namespace ExcelReaderWPF.Pages
{
    /// <summary>
    /// Interaction logic for CompareAllPages.xaml
    /// </summary>
    public partial class CompareAllPages : Page
    {
        public CompareAllPages()
        {
            InitializeComponent();
        }

		private static void FolderSelect(TextBox textbox)
		{
			var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
			var result = folderDialog.ShowDialog();
			switch (result)
			{
				case System.Windows.Forms.DialogResult.OK:
					var file = folderDialog.SelectedPath;
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
			FolderSelect(Folder1Textbox);
		}

		private void Folder2Select_Click(object sender, RoutedEventArgs e)
		{
			FolderSelect(Folder2Textbox);
		}

		private void File3Select_Click(object sender, RoutedEventArgs e)
		{
			var fileDialog = new System.Windows.Forms.OpenFileDialog();
			fileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
			var result = fileDialog.ShowDialog();
			switch (result)
			{
				case System.Windows.Forms.DialogResult.OK:
					var file = fileDialog.FileName;
					FileOutTextbox.Text = file;
					FileOutTextbox.ToolTip = file;
					break;
				case System.Windows.Forms.DialogResult.Cancel:
				default:
					FileOutTextbox.Text = null;
					FileOutTextbox.ToolTip = null;
					break;
			}
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			if(Folder1Textbox.Text != "" && Folder2Textbox.Text != "" && FileOutTextbox.Text != "" &&Folder1Textbox.Text != Folder2Textbox.Text)
			{
				if(CS_Files.Verification.VerifyFolders(Folder1Textbox.Text, Folder2Textbox.Text))
				{
					string[] firstAssess = Directory.GetFiles(Folder1Textbox.Text, "*.xlsx")
						.Select(System.IO.Path.GetFileNameWithoutExtension)
						.Select(p => p.Substring(0))
						.ToArray();
					string[] secondAssess = Directory.GetFiles(Folder2Textbox.Text, "*.xlsx")
						.Select(System.IO.Path.GetFileNameWithoutExtension)
						.Select(p => p.Substring(0))
						.ToArray();
					int index = 1;
					foreach (string file in firstAssess)//Currently tries to do every single foreach at the same time, it should do one after another
					{
						string file1 = Folder1Textbox.Text + "\\" + file + ".xlsx";
						string file2 = Folder2Textbox.Text + "\\" + (secondAssess[Array.IndexOf(secondAssess, file + " Second")]) + ".xlsx";

						//Might want to just create the progressbar and do the foreach within it, rather than doing a foreach generating progressbar objects
						//Pass the info into the progressbar, make a second constructor to take different information
						ProgressBar p = new ProgressBar(file1, file2, FileOutTextbox.Text, index);
						this.NavigationService.Navigate(p);
					}
				}
				else
				{
					MessageBoxResult foldersDontMatch = MessageBox.Show("The two selected folders do not have the correct naming format. Refer to the user guide.", "Error", MessageBoxButton.OK);
				}
			}
		}
	}
}
