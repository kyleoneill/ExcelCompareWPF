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

namespace ExcelReaderWPF.Pages
{
    /// <summary>
    /// Interaction logic for FirstPage.xaml
    /// </summary>
    public partial class FirstPage : Page
    {
        public FirstPage()
        {
            InitializeComponent();
        }

		private void CompareAllButton_Click(object sender, RoutedEventArgs e)
		{
			this.NavigationService.Navigate(new Uri("Pages/CompareAllPage.xaml", UriKind.Relative));
		}

		private void CompareTwoButton_Click(object sender, RoutedEventArgs e)
		{
			this.NavigationService.Navigate(new Uri("Pages/CompareTwoPage.xaml", UriKind.Relative));
		}
	}
}
