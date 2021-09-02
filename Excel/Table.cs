using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    class Table
    {
        public Dictionary<string, List<string>> Items = new Dictionary<string, List<string>>();

        public List<string> Headers = new List<string>();

        private int OrderBy = 1;

        private ExcelApp excelApp;

        public void SetExcelApp(ExcelApp excelApp)
        {
            this.excelApp = excelApp;
        }

        private List<int> valueIndex = new List<int>();

        public void AddHeaders(int x1, int y1, int x2, int y2)
        {
            bool check = true;

            for (int i = x1; i <= x2; i++)
            {
                for (int j = y1; j <= y2; j++)
                {
                    string cell = excelApp.ReadCell(i, j);

                    if (cell != "")
                    {
                        
                        Headers.Add(cell);

                        // list of index of values. this is to avoid the error of an empty string in the value cell
                        if (check)
                        {
                            valueIndex.Add(j);
                        }
                    }
                }
                check = false;
            }
        }

        public void AddValues(int y1, int x2, int y2, int orderby = 1)
        {
            string key = "";

            List<string> temp = new List<string>();

            if (orderby > y2)
            {
                this.OrderBy = 1;
            }
            else
            {
                this.OrderBy = orderby;
            }

            int i = x2 + 1;

            string cell = excelApp.ReadCell(i, y1);

            // chech if new line is clear
            while (cell != "")
            {
                // set key
                for (int j = y1; j <= y2; j++)
                {
                    if (j == OrderBy)
                    {
                        key = excelApp.ReadCell(i, j);
                    }
                }
                // create temp list of values
                for (int j = y1; j <= y2; j++)
                {
                    cell = excelApp.ReadCell(i, j);

                    temp.Add(cell);

                }
                // find clear cells and write 0 to avoid errors
                foreach (int index in valueIndex)
                {
                    if (temp[index-y1] == "")
                    {
                        temp[index-y1] = "0";
                    }
                }
                // add digital values of items with same key
                if (Items.ContainsKey(key))
                {
                    for (int k = 0; k < Items[key].Count; k++)
                    {
                        if (double.TryParse(temp[k], out double x) && double.TryParse(Items[key][k], out double y) && k != OrderBy)
                        {
                            string item = (y + x).ToString();

                            Items[key][k] = item;
                        }
                    }
                }
                else
                {
                    Items.Add(key, temp);
                }

                cell = excelApp.ReadCell(i + 1, 1);

                i++;

                temp = new List<string>();
            }
        }
    }
}
