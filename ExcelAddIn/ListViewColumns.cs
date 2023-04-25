using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn
{
    class ListViewColumns
    {
        string columnHeaders;
        int columnWeight;

        public ListViewColumns(string columnHeaders, int columnWeight)
        {
            this.columnHeaders = columnHeaders;
            this.columnWeight = columnWeight;
        }

        public string getColumnHeaders()
        {
            return columnHeaders;
        }

        public int getColumnWeight()
        {
            return columnWeight;
        }
    }
}
