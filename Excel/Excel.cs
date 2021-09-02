using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Excel
{
    class ExcelApp
    {
        Application excel = new _Excel.Application();

        string path;

        Workbook wb;

        Worksheet ws;

        Dictionary<string, List<string>> columns = new Dictionary<string, List<string>>();

        public ExcelApp(string path, int sheet)
        {
            this.path = path;

            wb = excel.Workbooks.Open(path);

            ws = wb.Worksheets[sheet];
        }

        public void TableToExcelFile(Table table)
        {
            Workbook workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            Worksheet worksheet = workbook.Worksheets[1];

            int i = 2, j = 1;

            // set headers

            foreach(var header in table.Headers)
            {
                worksheet.Cells[i-1, j].Value = header;

                j++;
            }

            // set values

            foreach(var item in table.Items)
            {
                j = 1;
               foreach(var value in  item.Value)
                {
                    worksheet.Cells[i, j].Value = value;

                    j++;
                }
                i++;
            }

            workbook.SaveAs(@"E:\хз\excel\5.xls");

            workbook.Close();
        }

        public void Close()
        {
            wb.Close();
        }

        public string ReadCell(int i, int j)
        {

            if(ws.Cells[i,j].Value != null)
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
