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
            this.Loaded += new RoutedEventHandler(OnLoad);
        }
        public void OnLoad(Object sender, RoutedEventArgs e)
        {
            int writebookRow = 2;

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

            foreach(Excel.Worksheet sheet in book1.Worksheets)
            {
                Excel.Worksheet compareSheet = book2.Sheets[sheet.Index];
                Excel.Range compareRange = compareSheet.UsedRange;
                //writebookRow = CS_Files.CompareSheets.Compare(sheet, compareRange, writeRange, writebookRow, pBar);
                int writebookColumns = 1;
                bool written;
                Excel.Range range = sheet.UsedRange;
                int rows = range.Rows.Count;
                int columns = range.Columns.Count;
                for (int i = 1; i <= rows; i++)
                {
                    written = false;
                    float percentage = ((float)i / rows) * 100;
                    pBar.Value = percentage;

                    if (range.Cells[i, 4].Value == compareRange.Cells[i, 4].Value && range.Cells[i, 4].Value != null)
                    {
                        for (int j = 5; j <= 11; j++)
                        {
                            if (range.Cells[i, j].Value == compareRange.Cells[i, j].Value && range.Cells[i, j].Value != null)
                            {
                                if (dictionary.Contains(Convert.ToString(range.Cells[i, j].Value)) || int.TryParse(Convert.ToString(range.Cells[i, j].Value), out int n))
                                {
                                    break;
                                }
                                writeRange.Cells[writebookRow, 1].Value = sheet.Name;
                                writeRange.Cells[writebookRow, 2].Value = range.Cells[1, j].Value;
                                writeRange.Cells[writebookRow, 3].Value = range.Cells[i, 4].Value;
                                writeRange.Cells[writebookRow, writebookColumns + 3].Value = range.Cells[i, j].Value;
                                writebookColumns++;
                                written = true;
                            }
                        }
                        if (written)
                            writebookRow++;
                        writebookColumns = 1;
                    }
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            book1.Close();
            book2.Close();
            bookOut.Save();
            bookOut.Close();
            app.Quit();
        }
    }
}
