using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection.Emit;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Form_VisitsAddEntry : Form
    {
        Excel.Application application;
        Excel.Workbook activeWorkBook;
        Excel.Worksheet activeWorksheet;
        Range activeCell;
        int columnCount;
        int[] headersIndex;
        int rowCount;
        int rowToInsert;
        int activeCellRow;
        string[,] stringArray;

        public Form_VisitsAddEntry(
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

            rowToInsert = this.rowCount + 1;
            activeCellRow = this.activeCell.Row;

            stringArray = MyRibbon.GetDataToStringArray(application);

            // настраиваем заголоки listView
            foreach ((string, int) visitsHeaderAndWeight in MyRibbon.visitsHeadersAndWeights)
            {
                listView.Columns.Add(visitsHeaderAndWeight.Item1, visitsHeaderAndWeight.Item2);
            }

            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("ru-RU"));

            textBox.Select();
        }

        private void checkBox_toEndOfList_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_toEndOfList.Checked)
            {
                rowToInsert = rowCount + 1;
            }
            else
            {
                rowToInsert = activeCellRow;
            }
        }

        private void checkBox_findSurname_CheckedChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void checkBox_withoutDuplicatingEntrys_CheckedChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            button_addSelected.Enabled = listView.SelectedItems.Count > 0;
        }

        private void updateListview()
        {
            button_addSelected.Enabled = false;

            listView.Items.Clear();

            if (textBox.Text.Length >= 3)
            {                
                List<List<string[]>> collectionOfElements = MyRibbon.findEntryInStringArray(stringArray, textBox.Text, checkBox_findSurname.Checked, checkBox_withoutDuplicatingEntrys.Checked);

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
        }

        private void button_addNew_Click(object sender, EventArgs e)
        {
            activeWorksheet.Cells[rowToInsert, headersIndex[0]].value = textBox.Text;

            if (checkBox_addCurTime.Checked && headersIndex[0] - 1 > 0)
            {
                activeWorksheet.Cells[rowToInsert, headersIndex[0] - 1].value = DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
            }

            Close();
        }

        private void button_addSelected_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < headersIndex.Length; i++)
            {
                if (headersIndex[i] != 0)
                {
                    activeWorksheet.Cells[rowToInsert, headersIndex[i]].value = listView.SelectedItems[0].SubItems[i].Text;
                }
            }
            if (checkBox_addCurTime.Checked && headersIndex[0] - 1 >0)
            {
                activeWorksheet.Cells[rowToInsert, headersIndex[0] - 1].value = DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
            }

            Close();
        }
    }
}
