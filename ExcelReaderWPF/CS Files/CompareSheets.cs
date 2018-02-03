using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReaderWPF.CS_Files
{
    class CompareSheets
    {
        public static int Compare(Excel.Worksheet sheet, Excel.Workbook workbook2, Excel.Range writeRange, int writebookRow)
        {
            List<String> dictionary = new List<string>()
            {
                "CABLE BOX",
                "FACE-PLATE",
                "WIRE-MOLD",
                "CABLE TV",
                "DATA JACK",
                "PHONE JACK",
                "RED PHONE",
                "CABLE " + "\n" + "BOX ", //Replace these two with a regex attempt to find all cable box?
                "CABLE" + "\n" + "BOX ",
            };
            int writebookColumns = 1;
            bool written;
            Excel.Worksheet compareSheet = workbook2.Sheets[sheet.Index];
            Excel.Range compareRange = compareSheet.UsedRange;
            Excel.Range range = sheet.UsedRange;
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;
            for (int i = 1; i <= rows; i++)
            {
                written = false;
                float percentage = ((float)i / rows) * 100;
                //Console.Write(sheet.Name + ": " + percentage.ToString("0.0") + "%");
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
            return writebookRow;
        }
    }
}
