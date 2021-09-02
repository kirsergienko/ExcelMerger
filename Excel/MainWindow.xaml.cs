using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace Excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            ExcelApp excel = new ExcelApp(@"E:\хз\excel\1.xls", 1);

            Table table = new Table();

            table.SetExcelApp(excel);

            int x1 = 10, x2 = 10, y1 = 1, y2 = 7;

            table.AddHeaders(x1, y1, x2, y2);

            table.AddValues(y1, x2, y2);

            excel = new ExcelApp(@"E:\хз\excel\2.xls", 1);

            table.SetExcelApp(excel);

            table.AddValues(y1, x2, y2);

            excel = new ExcelApp(@"E:\хз\excel\3.xls", 1);

            table.SetExcelApp(excel);

            table.AddValues(y1, x2, y2);

            excel = new ExcelApp(@"E:\хз\excel\4.xls", 1);

            table.SetExcelApp(excel);

            table.AddValues(y1, x2, y2);

            excel.TableToExcelFile(table);

            excel.Close();
        }
    }
}
