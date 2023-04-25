using ExcelAddIn;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Form_InterestsEntrySearch : Form
    {
        Excel.Application application;
        Excel.Workbook activeWorkBook;
        Excel.Worksheet activeWorksheet;
        Range activeCell;
        int[] columnIndex;
        string[] personNames;

        public Form_InterestsEntrySearch(
            Excel.Application application,
            Excel.Workbook activeWorkBook,
            Excel.Worksheet activeWorksheet,
            Range activeCell,
            int[] columnIndex,
            string[] personNames
        )
        {
            InitializeComponent();

            this.application = application;
            this.activeWorkBook = activeWorkBook;
            this.activeWorksheet = activeWorksheet;
            this.activeCell = activeCell;
            this.columnIndex = columnIndex;
            this.personNames = personNames;

            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("ru-RU"));

            textBox.Select();
        }

        private void checkBox_findSurname_CheckedChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            updateListview();
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            button_EditChanged.Enabled = listView.SelectedItems.Count > 0;

            if (listView.SelectedItems.Count > 0)
            {
                activeWorksheet.Rows[int.Parse(listView.SelectedItems[0].SubItems[1].Text)].Select();
            }
        }

        private void updateListview()
        {
            button_EditChanged.Enabled = false;

            listView.Items.Clear();

            if (textBox.Text.Length >= 3)
            {
                List<List<string[]>> listOfList = findEntry(checkBox_findSurname.Checked, textBox.Text);

                var items = new List<ListViewItem>();
                foreach (List<string[]> listOflElements in listOfList)
                {
                    foreach (string[] element in listOflElements)
                    {
                        items.Add(new ListViewItem(element, -1));
                    }
                }

                listView.BeginUpdate();
                listView.Items.AddRange(items.ToArray());
                listView.EndUpdate();
            }
        }

        private void button_EditChanged_Click(object sender, EventArgs e)
        {
            new Form_InterestsEntryEditor(
                application,
                activeWorkBook,
                activeWorksheet,
                activeCell,
                columnIndex,
                personNames,
                int.Parse(listView.SelectedItems[0].SubItems[1].Text)
            ).Show();

            Close();
        }

        private void button_CreateNew_Click(object sender, EventArgs e)
        {
            new Form_InterestsEntryEditor(
                application,
                activeWorkBook,
                activeWorksheet,
                activeCell,
                columnIndex,
                personNames,
                1
            ).Show();

            Close();
        }


        private List<List<string[]>> findEntry(bool insertSurname, string textToFinding)
        {
            List<List<string[]>> listOfList = new List<List<string[]>>();
            for (int i = 0; i < 8; i++) listOfList.Add(new List<string[]>());

            int curItem;
            string personName;

            for (int i = 2; i < personNames.Length; i++)
            {
                personName = personNames[i];

                if (string.IsNullOrEmpty(personName)) continue;

                curItem = -1;
                //точное совпадение
                if (string.Equals(personName, textToFinding, StringComparison.OrdinalIgnoreCase)) curItem = 0;
                //точное совпадение с фамилией 
                else if (string.Equals(personName.Split(' ')[0], textToFinding, StringComparison.OrdinalIgnoreCase)) curItem = 1;
                //точное совпадение фамилии с фамилией
                else if (insertSurname && string.Equals(personName.Split(' ')[0], textToFinding.Split(' ')[0], StringComparison.OrdinalIgnoreCase)) curItem = 2;
                //точное совпадение с началом фамилии
                else if (personName.Length >= textToFinding.Length && string.Equals(personName.Substring(0, textToFinding.Length), textToFinding, StringComparison.OrdinalIgnoreCase)) curItem = 3;
                //вхождение в фамилию       
                else if (personName.Split(' ')[0].IndexOf(textToFinding, StringComparison.OrdinalIgnoreCase) >= 0) curItem = 4;
                //вхождение фамилии в фамилию
                else if (insertSurname && personName.Split(' ')[0].IndexOf(textToFinding.Split(' ')[0], StringComparison.OrdinalIgnoreCase) >= 0) curItem = 5;
                //вхождение в ФИО
                else if (personName.IndexOf(textToFinding, StringComparison.OrdinalIgnoreCase) >= 0) curItem = 6;
                //вхождение фамилии в ФИО
                else if (insertSurname && personName.IndexOf(textToFinding.Split(' ')[0], StringComparison.OrdinalIgnoreCase) >= 0) curItem = 7;

                if (curItem >= 0)
                {
                    listOfList[curItem].Add(new string[] { personName, (i + 1).ToString() });
                }
            }

            return listOfList;
        }

    }
}
