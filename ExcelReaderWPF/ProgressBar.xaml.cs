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
        List<String> dictionary = new List<string>()
            {
                "CABLE BOX",
                "FACE-PLATE",
                "WIRE-MOLD",
                "CABLE TV",
                "DATA JACK",
                "PHONE JACK",
                "RED PHONE",
                "CABLE " + "\n" + "BOX "
            };
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
        /*public void OnLoad(Object sender, RoutedEventArgs e)
        {

        }*/
        
        public static void UpdateProgressBar()
        {
            ProgressBarObj.Dispatcher.Invoke();
            //The heck ^?
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            Excel.Application app = new Excel.Application();
            Excel.Workbook book1 = app.Workbooks.Open(File1Path);
            Excel.Workbook book2 = app.Workbooks.Open(File2Path);
            Excel.Workbook bookOut = app.Workbooks.Open(FileOutPath);

            Excel.Worksheet writeSheet = bookOut.Sheets[1];
            Excel.Range writeRange = writeSheet.UsedRange;

            writeSheet.Name = book1.Name;
            writeRange.Cells[1, 1].Value = "Hall";
            writeRange.Cells[1, 2].Value = "Jack Type";
            writeRange.Cells[1, 3].Value = "Jack";

            foreach (Excel.Worksheet sheet in book1.Worksheets)
            {
                //Insert a method in this file that updates the progress bar when called. Call this method in the CompareSheets class, compare method.
                //make sure to pass an INT. Convert float to int before passing.
            }
        }    }
}
