using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

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
					ProgressBar p = new ProgressBar(Folder1Textbox.Text, Folder2Textbox.Text, FileOutTextbox.Text);
					this.NavigationService.Navigate(p);
				}
				else
				{
					MessageBoxResult foldersDontMatch = MessageBox.Show("The two selected folders do not have the correct naming format. Refer to the user guide.", "Error", MessageBoxButton.OK);
				}
			}
		}
	}
}
