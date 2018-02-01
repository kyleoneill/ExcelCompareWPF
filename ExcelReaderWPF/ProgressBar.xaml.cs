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

namespace ExcelReaderWPF
{
    /// <summary>
    /// Interaction logic for ProgressBar.xaml
    /// </summary>
    public partial class ProgressBar : Page
    {
        public string file1Path { get; set; }
        public string file2Path { get; set; }
        public string fileOutPath { get; set; }
        public int fileOutIndex { get; set; }
        public ProgressBar()
        {
            InitializeComponent();
        }
    }
}
