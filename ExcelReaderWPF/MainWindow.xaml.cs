using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.ComponentModel;

namespace ExcelReaderWPF
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : NavigationWindow
	{
		public MainWindow()
		{
			InitializeComponent();
		}

        void DataWindow_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                foreach (Process proc in Process.GetProcessesByName("EXCEL"))
                {
                    proc.Kill();
                }
            }
            //catch (Exception ex)
            catch
            {
                MessageBoxResult excelKill = System.Windows.MessageBox.Show("Failure to close all Excel processes.", "Error", MessageBoxButton.OK);
            }
        }
	}
}
