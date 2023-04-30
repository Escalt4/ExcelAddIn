using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Form_TyposCorrection : Form
    {
        Excel.Application application;
        Excel.Workbook activeWorkBook;
        Excel.Worksheet activeWorksheet;
        Range activeCell;
        int curColumn;
        Dictionary<string, List<int>> stringDictionary = new Dictionary<string, List<int>>();
        List<string> stringDictionaryKeys;

        public Form_TyposCorrection(Excel.Application application, Excel.Workbook activeWorkBook, Excel.Worksheet activeWorksheet, Range activeCell)
        {
            InitializeComponent();

            this.application = application;
            this.activeWorkBook = activeWorkBook;
            this.activeWorksheet = activeWorksheet;
            this.activeCell = activeCell;

            curColumn = activeCell.Column;
            
            int rowCount = MyRibbon.getRowCount(activeWorksheet, curColumn);
            if (rowCount <= 1) rowCount = 2;

            object[,] curRange = activeWorksheet.Range[activeWorksheet.Cells[1, curColumn], activeWorksheet.Cells[rowCount, curColumn]].value2;
            for (int i = 2; i <= rowCount; i++)
            {
                if (curRange[i, 1] != null && curRange[i, 1] is string)
                {
                    string key = curRange[i, 1].ToString();
                    if (stringDictionary.ContainsKey(key))
                    {
                        stringDictionary[key].Add(i);
                    }
                    else
                    {
                        stringDictionary.Add(key, new List<int>());
                        stringDictionary[key].Add(i);
                    }
                }                
            }

            stringDictionaryKeys = stringDictionary.Keys.ToList();

            for (int i = 0; i < stringDictionaryKeys.Count; i++)
            {
                dataGridView.Rows.Add(i,stringDictionaryKeys[i]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dataGridViewRow in dataGridView.Rows)
            {
                Debug.Write(dataGridViewRow.Cells[0].Value);
                Debug.Write(" ");
                Debug.Write(dataGridViewRow.Cells[1].Value);
                Debug.Write("\n");
            } 
        }                
    }
}
