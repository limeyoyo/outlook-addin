using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MyOutlookAddIn
{
    class ExcelManage
    {
        private Application application = null;
        private Workbook wBook = null;
        private Worksheet wSheet = null;
        private Range allColumn = null;
        private static int row = 0;

        public const string Alphabet = "ABCDEFGHIGKLMNOPQRSTUVWXYZ";

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);

        public ExcelManage()
        {
            this.application = new Application();
            this.application.DisplayAlerts = false;
            this.application.AlertBeforeOverwriting = false;
            this.application.Visible = false;
            this.wBook = application.Workbooks.Add(Type.Missing);
            this.wSheet = (Worksheet)wBook.ActiveSheet;
            this.allColumn = wSheet.Columns;
            row = 1;
        }

        public void WriteToSheet(List<string> str)
        {
            for (int i = 0; i < str.Count; i++)
            {
                if (row > 1 && 0 == i)
                {
                    wSheet.get_Range((string)(Alphabet[i] + "") + row, Type.Missing).Value2 = row - 1;
                }
                else
                {
                    wSheet.get_Range((string)(Alphabet[i] + "") + row, Type.Missing).Value2 = str[i];
                }
            }
            row += 1;
        }

        public void SetFormat()
        {
            this.allColumn["A:A", System.Type.Missing].ColumnWidth = 5;
            this.allColumn["B:B", System.Type.Missing].ColumnWidth = 18;
            this.allColumn["C:C", System.Type.Missing].ColumnWidth = 30;
            this.allColumn["D:D", System.Type.Missing].ColumnWidth = 30;
            this.allColumn["E:E", System.Type.Missing].ColumnWidth = 30;
            this.allColumn["F:F", System.Type.Missing].ColumnWidth = 40;
            this.allColumn.WrapText = true;
        }

        public void SaveExcel(String path)
        {
            wBook.SaveAs(path, Type.Missing, "", "", Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);
        }

        public void Close()
        {
            IntPtr intptr = new IntPtr(this.application.Hwnd);
            int id;
            GetWindowThreadProcessId(intptr, out id);
            this.wBook.Close();
            this.application.Quit();
            if (this.wBook != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.wBook);
            }
            if (this.wBook != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.application);
            }
            Process.GetProcessById(id).Kill();
        }
    }
}
