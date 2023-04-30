using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Form_VisitsAutocomplete : Form
    {
        Excel.Application application;
        Excel.Workbook activeWorkBook;
        Excel.Worksheet activeWorksheet;
        Range activeCell;
        int columnCount;
        int[] headersIndex;
        int rowCount;
        int activeCellRow;
        string activeCellText;
        string[,] stringArray;

        public Form_VisitsAutocomplete(
            Excel.Application application,
            Excel.Workbook activeWorkBook,
            Excel.Worksheet activeWorksheet,
            Range activeCell,
            int columnCount,
            int[] headersIndex,
            int rowCount
        )
        {
            InitializeComponent();

            this.application = application;
            this.activeWorkBook = activeWorkBook;
            this.activeWorksheet = activeWorksheet;
            this.activeCell = activeCell;
            this.columnCount = columnCount;
            this.headersIndex = headersIndex;
            this.rowCount = rowCount;

            activeCellRow = this.activeCell.Row;
            activeCellText = this.activeCell.Text.Trim();

            stringArray = MyRibbon.GetDataToStringArray(application, true);

            // настраиваем заголоки listView
            foreach ((string, int) visitsHeaderAndWeight in MyRibbon.visitsHeadersAndWeights)
            {
                listView.Columns.Add(visitsHeaderAndWeight.Item1, visitsHeaderAndWeight.Item2);
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            updateListview();
        }

        private void checkBox_findSurname_CheckedChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void checkBox_withoutDuplicatingEntrys_CheckedChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void updateListview()
        {
            button_addSelected.Enabled = false;

            listView.Items.Clear();

            List<List<string[]>> collectionOfElements = MyRibbon.findEntryInStringArray(stringArray, activeCellText, checkBox_findSurname.Checked, checkBox_withoutDuplicatingEntrys.Checked);

            var items = new List<ListViewItem>();
            foreach (List<string[]> collElements in collectionOfElements)
            {
                foreach (string[] element in collElements)
                {
                    items.Add(new ListViewItem(element, -1));
                }
            }

            listView.BeginUpdate();
            listView.Items.AddRange(items.ToArray());
            listView.EndUpdate();
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            button_addSelected.Enabled = listView.SelectedItems.Count > 0;
        }

        private void button_addSelected_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < headersIndex.Length; i++)
            {
                if (headersIndex[i] != 0)
                {
                    activeWorksheet.Cells[activeCellRow, headersIndex[i]].value = listView.SelectedItems[0].SubItems[i].Text;
                }
            }

            Close();
        }
    }
}
