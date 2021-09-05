using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using Microsoft.Win32;
using System.Threading;

namespace Excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        int r1 = 0, r2 = 0, c2 = 0, c1 = 0;

        int orderby = 1;

        string path;

        bool sameTables = true;

        Table table = new Table();

        Thread thread;

        List<ExcelApp> excelApp = new List<ExcelApp>();

        ExcelApp newApp;

        public MainWindow()
        {
            InitializeComponent();

            this.Closed += MainWindow_Closed;
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            if (thread != null)
            {
                thread.Abort();
            }

            foreach (var app in excelApp)
            {
                if (app != null)
                {
                    app.Close();

                    if (newApp != null)
                    {
                        newApp.Close();

                    }
                }
            }
        }

        private void addButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files(.xls)|*.xls|  Excel Files(.xlsx)| *.xlsx | Excel Files(*.xlsm) | *.xlsm";

            if (openFileDialog.ShowDialog() == true)
            {
                listbox.Items.Add(openFileDialog.FileName);
            }
        }

        private void removeButton_Click(object sender, RoutedEventArgs e)
        {
            listbox.Items.Remove(listbox.SelectedItem);
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            saveFileDialog.FileName = "NewExcelFile";

            saveFileDialog.DefaultExt = ".xls";

            var tokenSource = new CancellationTokenSource();

            var token = tokenSource.Token;

            if (saveFileDialog.ShowDialog() == true)
            {
                CreateExcelFiles();

                newApp = new ExcelApp();

                path = saveFileDialog.FileName;

                ThreadStart threadst = new ThreadStart(() => { Start(); });

                thread = new Thread(threadst);

                thread.IsBackground = true;

                thread.Start();

                loadingStackPanel.Visibility = Visibility.Visible;

            }
        }

        private void LoadEnd()
        {
            loadingStackPanel.Dispatcher.Invoke(() =>
            {
                loadingStackPanel.Visibility = Visibility.Hidden;
            });
        }

        private void CreateExcelFiles()
        {
            foreach (var file in listbox.Items)
            {
                var app = new ExcelApp(file.ToString(), 1);

                app.ExcelClosed += App_ExcelClosed;

                excelApp.Add(app);
            }
        }

        private void App_ExcelClosed(object o)
        {
            (o as ExcelApp).Opened = false;
        }

        public void Start()
        {
            if (CheckningSettings())
            {
                table.SetExcelApp(excelApp[0]);

                table.AddHeaders(r1, c1, r2, c2);

                // set value 

                int i = 0;

                foreach (var app in excelApp)
                {
                    if (i != 0)
                    {
                        table.SetExcelApp(app);
                    }

                    try
                    {
                        if (excelApp[i].Opened)
                        {
                            table.AddValues(c1, r2, c2, orderby);
                        }
                    }
                    catch (Exception ex)
                    {
                        sameTables = false;
                    }
                    i++;
                }
                // fill excel file
                if (sameTables)
                {
                    excelApp[0].TableToExcelFile(table, path);
                }
                else
                {
                    MessageBox.Show("Таблицы не одинаковы!");
                }

                LoadEnd();

            }
        }


        public bool CheckningSettings()
        {
            bool check = false;

            Dispatcher.Invoke(new Action(() =>
            {
                if (listbox.Items.Count > 0)
                {
                    if (int.TryParse(r1textbox.Text, out r1) && int.TryParse(c1textbox.Text, out c1) && int.TryParse(c2textbox.Text, out c2)
                         && c2 > c1)
                    {
                        if (int.TryParse(orderByTextbox.Text, out orderby) && orderby <= c2)
                        {
                            check = true;

                            r2 = r1;
                        }
                        else
                        {
                            MessageBox.Show("Номер столбца заполнен неправильно!");
                            check = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Шапка заполнена неправильно!");
                        check = false;
                    }
                }
                else
                {
                    MessageBox.Show("Файлы не выбраны!");
                    check = false;
                }
            }));

            return check;
        }

    }
}
