using ExcelAddIn;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Form_InterestsEntryEditor : Form
    {
        Excel.Application application;
        Excel.Workbook activeWorkBook;
        Excel.Worksheet activeWorksheet;
        Range activeCell;
        int[] columnIndex;
        string[] personNames;
        int changedRow;

        public Form_InterestsEntryEditor(
            Excel.Application application,
            Excel.Workbook activeWorkBook,
            Excel.Worksheet activeWorksheet,
            Range activeCell,
            int[] columnIndex,
            string[] personNames,
            int changedRow
        )
        {
            InitializeComponent();

            this.application = application;
            this.activeWorkBook = activeWorkBook;
            this.activeWorksheet = activeWorksheet;
            this.activeCell = activeCell;
            this.columnIndex = columnIndex;
            this.personNames = personNames;
            this.changedRow = changedRow;

            if (changedRow != 1)
            {
                button.Text = "Применить";
                textBox_PersonName.Text = activeWorksheet.Cells[changedRow, 2].text;
                textBox_PhoneNumber.Text = activeWorksheet.Cells[changedRow, 3].text;
            }

            // получаем список интересов
            List<(string, bool)> listOfInterests = new List<(string, bool)>();
            int curColunm = 4;
            string curWorksheetCells = activeWorksheet.Cells[2, curColunm].text;
            while (curWorksheetCells != "")
            {

                listOfInterests.Add((curWorksheetCells, activeWorksheet.Cells[changedRow, curColunm].text == "+"));

                curColunm++;
                curWorksheetCells = activeWorksheet.Cells[2, curColunm].text;
            }

            listView.BeginUpdate();
            Color backColor;
            foreach ((string, bool) interest in listOfInterests)
            {
                if (interest.Item2) backColor = Color.PaleGreen;
                else backColor = Color.White;

                listView.Items.Add(new ListViewItem(interest.Item1, -1) { Checked = interest.Item2, BackColor = backColor });
            }
            listView.EndUpdate();
        }

        private void listView_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (listView.Items[e.Index].Checked)
            {
                listView.Items[e.Index].BackColor = Color.White;
            }
            else
            {
                listView.Items[e.Index].BackColor = Color.PaleGreen;
            }
            listView.Items[e.Index].Selected = false;
        }


        private void button_Click(object sender, EventArgs e)
        {
            int rowToInsert;

            if (changedRow != 1)
            {
                rowToInsert = changedRow;
            }
            else
            {
                application.CutCopyMode = 0;
                rowToInsert = 3;
                for (; rowToInsert <= personNames.Length; rowToInsert++)
                {
                    if (string.Compare(textBox_PersonName.Text, personNames[rowToInsert - 1]) < 0)
                    {
                        activeWorksheet.Rows[rowToInsert].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        activeWorksheet.Rows[rowToInsert + 1].Copy();

                        goto skip1;
                    }
                }
                activeWorksheet.Rows[rowToInsert].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                activeWorksheet.Rows[rowToInsert - 1].Copy();
            skip1:
                activeWorksheet.Rows[rowToInsert].PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                application.CutCopyMode = 0;
            }

            activeWorksheet.Cells[rowToInsert, 2].value = textBox_PersonName.Text;
            activeWorksheet.Cells[rowToInsert, 3].value = textBox_PhoneNumber.Text;

            for (int i = 0; i < listView.Items.Count; i++)
            {
                if (listView.Items[i].Checked)
                {
                    activeWorksheet.Cells[rowToInsert, i + 4].value = "+";
                }
                else
                {
                    activeWorksheet.Cells[rowToInsert, i + 4].value = "";
                }
            }

            activeWorksheet.Rows[rowToInsert].Select();

            Close();
        }
    }
}
