using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Excel
{
    public class ExcelApp
    {
        public bool Opened { get; set; }

        Application excel = new _Excel.Application();

        string path;

        Workbook wb;

        Worksheet ws;

        Dictionary<string, List<string>> columns = new Dictionary<string, List<string>>();

        public delegate void EventHandler(object o);

        public event EventHandler ExcelClosed;

        protected virtual void Closed()
        {
            ExcelClosed?.Invoke(this);
        }

        public ExcelApp()
        {

        }

        public ExcelApp(string path, int sheet)
        {
            this.path = path;

            wb = excel.Workbooks.Open(path);

            ws = wb.Worksheets[sheet];

            Opened = true;

        }

        public void TableToExcelFile(Table table, string path)
        {
            wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            ws = wb.Worksheets[1];

            int i = 2, j = 1;

            // set headers

            foreach (var header in table.Headers)
            {
                ws.Cells[i - 1, j].Value = header;

                j++;
            }

            // set values

            foreach (var item in table.Items)
            {
                j = 1;
                foreach (var value in item.Value)
                {
                    ws.Cells[i, j].Value = value;

                    j++;
                }
                i++;
            }

            wb.SaveAs(path);
        }

        public void Close()
        {
            Opened = false;

            Closed();

            if (wb != null)
            {
                wb.Close();
            }

            if (excel != null)
            {
                excel.Quit();
            }

            try
            {
                Marshal.FinalReleaseComObject(excel);
                Marshal.FinalReleaseComObject(ws);
                Marshal.FinalReleaseComObject(wb);
            }
            catch (Exception ex)
            {

            }
            finally
            {
                ws = null;
                wb = null;
                excel = null;
            }

        }

        public string ReadCell(int i, int j)
        {
            if (ws.Cells[i, j].Value != null && Opened)
            {
                return ws.Cells[i, j].Value.ToString();
            }
            else
            {
                return "";
            }
        }
    }
}
