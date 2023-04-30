using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Office.Tools.Excel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

// TODO:  Выполните эти шаги, чтобы активировать элемент XML ленты:

// 1: Скопируйте следующий блок кода в класс ThisAddin, ThisWorkbook или ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Создайте методы обратного вызова в области "Обратные вызовы ленты" этого класса, чтобы обрабатывать действия
//    пользователя, например нажатие кнопки. Примечание: если эта лента экспортирована из конструктора ленты,
//    переместите свой код из обработчиков событий в методы обратного вызова и модифицируйте этот код, чтобы работать с
//    моделью программирования расширения ленты (RibbonX).

// 3. Назначьте атрибуты тегам элементов управления в XML-файле ленты, чтобы идентифицировать соответствующие методы обратного вызова в своем коде.  

// Дополнительные сведения можно найти в XML-документации для ленты в справке набора средств Visual Studio для Office.


namespace ExcelAddIn
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        // заголовки и ширина колонок 
        public static List<(string, int)> visitsHeadersAndWeights = new List<(string, int)>()
        {
             ("ФИО", 270) ,
             ("Пол", 40) ,
             ("Дата рождения*", 100) ,
             ("Контактный телефон*", 130) ,
             ("Район проживания", 115),
             ("Участник проекта «Московское долголетие»*", 50) ,
             ("Взято из файла файла", 220)
        };
        public static string[] visitsHeaders = getHeadersArray(visitsHeadersAndWeights, true);
        public static string[] birthdayHeaders = new string[] { "ФИО", "Дата рождения", "Контактный телефон", "Район проживания" };
        public static string[] makeIdenticalHeaders = new string[] { "Дата посещения", "ФИО", "Участник проекта «Московское долголетие»*" };

        public static string[] monthNameList = new string[] { "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
                                                        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь" };

        Form_VisitsAddEntry form_VisitsAddEntry;
        Form_InterestsEntrySearch form_InterestsEntrySearch;

        bool enableVisits = false;
        bool enableInterests = false;

        private Office.IRibbonUI ribbon;

        public MyRibbon() { }

        #region Элементы IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn.MyRibbon.xml");
        }

        #endregion

        #region Обратные вызовы ленты
        // Информацию о методах создания обратного вызова см. здесь. Дополнительные сведения о методах добавления обратного вызова см. по ссылке https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            Globals.ThisAddIn.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(WorkbookActivateHandler);
        }

        public bool GetVisitsEnabled(IRibbonControl control)
        {
            return enableVisits;
        }

        public bool GetInterestsEnabled(IRibbonControl control)
        {
            return enableInterests;
        }

        public void WorkbookActivateHandler(Excel.Workbook workbook)
        {
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            enableVisits = activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) >= 0;
            enableInterests = activeWorkBook.Name.IndexOf(Properties.Settings.Default.InterestsFileName, StringComparison.OrdinalIgnoreCase) >= 0;

            this.ribbon.Invalidate();
        }

        // добавить текущее время в активную ячейку
        public void VisitsAddCurTime(Office.IRibbonControl control)
        {
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            if (activeWorksheet.Cells[1, activeCell.Column].Text != "Время прихода") return;

            activeCell.Value = DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString();
        }

        // автокомплит выбраной записи в ведомости посещаемости
        public void VisitsEntryAutocomplete(Office.IRibbonControl control)
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            if (string.IsNullOrEmpty(activeCell.Text.Trim())) return;

            // получить количество стобцов с заголовками 
            int columnCount = MyRibbon.getColumnCount(activeWorksheet, 1);
            if (columnCount == 0) return;
            // получить индексы колонок по их заголовкам
            int[] headersIndex = MyRibbon.getHeadersIndex(activeWorksheet, visitsHeaders, 1);
            // если нет колонки ФИО пропускаем итерацию
            if (headersIndex[0] == 0) return;
            // получить количество строк в определенной колонке
            int rowCount = MyRibbon.getRowCount(activeWorksheet, headersIndex[0]);
            // если нет строк ФИО пропускаем итерацию
            if (rowCount == 0) return;

            if (activeWorksheet.Cells[1, activeCell.Column].Text == visitsHeaders[0])
            {
                new Form_VisitsAutocomplete(application, activeWorkBook, activeWorksheet, activeCell, columnCount, headersIndex, rowCount).ShowDialog();
            }
        }

        // добавить новую запись в ведомость посещаемости
        public void VisitsAddNewEntry(Office.IRibbonControl control)
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            // получить количество стобцов с заголовками 
            int columnCount = MyRibbon.getColumnCount(activeWorksheet, 1);
            if (columnCount == 0) return;
            // получить индексы колонок по их заголовкам
            int[] headersIndex = MyRibbon.getHeadersIndex(activeWorksheet, visitsHeaders, 1);
            // если нет колонки ФИО пропускаем итерацию
            if (headersIndex[0] == 0) return;
            // получить количество строк в определенной колонке
            int rowCount = MyRibbon.getRowCount(activeWorksheet, headersIndex[0]);
            // если нет строк ФИО пропускаем итерацию
            if (rowCount == 0) return;

            if (form_VisitsAddEntry == null || form_VisitsAddEntry.IsDisposed)
            {
                form_VisitsAddEntry = new Form_VisitsAddEntry(application, activeWorkBook, activeWorksheet, activeCell, columnCount, headersIndex, rowCount);
                form_VisitsAddEntry.Show();
            }
            else
            {
                form_VisitsAddEntry.WindowState = FormWindowState.Normal;
                form_VisitsAddEntry.Activate();
            }
        }

        // перейти к последней записи в ведомости посещаемости 
        public void VisitsToLastEntry(Office.IRibbonControl control)
        {
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            // получить количество стобцов с заголовками 
            int columnCount = MyRibbon.getColumnCount(activeWorksheet, 1);

            // если нет колонки ФИО пропускаем итерацию
            if (columnCount == 0) return;

            // получить индексы колонок по их заголовкам
            int[] headersIndex = MyRibbon.getHeadersIndex(activeWorksheet, visitsHeaders, 1);

            // если нет колонки ФИО пропускаем итерацию
            if (headersIndex[0] == 0) return;

            // получить количество строк в определенной колонке
            int rowCount = MyRibbon.getRowCount(activeWorksheet, headersIndex[0]);

            activeWorksheet.Rows[rowCount].Select();
        }

        // добавить человека из ведомости посещаемости в дни рождения
        public void VisitsAddToBirthday(Office.IRibbonControl control)
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            string birthdayWorkBookName = Properties.Settings.Default.BirthdayFileName + ".xlsx";

            //ищем индексы колонок в текущей книге 
            int[] activeWorksheetColumnIndex = getHeadersIndex(activeWorksheet, birthdayHeaders, 1, false);

            if (activeWorksheetColumnIndex[0] == 0 || activeWorksheetColumnIndex[1] == 0) return;

            // выход если ячейка в колонке ФИО или Дата рождения пустая
            if (activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[0]].Text == "" || activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[1]].Text == "")
            {
                return;
            }

            DateTime date;
            // выход если дату рождения нельзя распознать
            if (!DateTime.TryParse(activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[1]].Text, out date))
            {
                MessageBox.Show("Не удалось распознать дату рождения", "Ошибка!", 0, MessageBoxIcon.Error);
                return;
            }

            // среди открытых книг ищем книгу дней рождений
            Excel.Workbook birthdayWorkBook;
            try
            {
                birthdayWorkBook = application.Workbooks[birthdayWorkBookName];
            }
            catch (Exception)
            {
                MessageBox.Show("Книга " + birthdayWorkBookName + " не открыта", "Ошибка!", 0, MessageBoxIcon.Error);
                return;
            }

            // получаем лист месяца даты рождения
            Excel.Worksheet birthdayWorksheet;
            try
            {
                birthdayWorksheet = birthdayWorkBook.Sheets[monthNameList[date.Month]];
            }
            catch (Exception)
            {
                return;
            }

            // проверяем что в книге нет выбраного человека            
            foreach (Excel.Worksheet worksheet in birthdayWorkBook.Worksheets)
            {
                int[] headersIndex = getHeadersIndex(worksheet, birthdayHeaders, 1, false);
                // если нет колонки ФИО пропускаем итерацию
                if (headersIndex[0] == 0) continue;

                // получить количество строк в определенной колонке
                int rowCount = MyRibbon.getRowCount(worksheet, headersIndex[0]);
                if (rowCount <= 1) continue;

                // получаем диапазон 
                object[,] curRange = worksheet.Range[worksheet.Cells[1, headersIndex[0]], worksheet.Cells[rowCount, headersIndex[0]]].Value2;

                // перебираем все записи 
                for (int i = 2; i <= rowCount; i++)
                {
                    if (curRange[i, 1] is null) continue;

                    if (curRange[i, 1].ToString().Trim() == activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[0]].Text.Trim())
                    {
                        MessageBox.Show("Запись уже существует\n\nЛист:\t" + worksheet.Name + "\nСтрока:\t" + i, "Ошибка!", 0, MessageBoxIcon.Error);

                        return;
                    }
                }
            }

            //ищем индексы колонок в книге др
            int[] birthdayColumnIndex = getHeadersIndex(birthdayWorksheet, birthdayHeaders, 1, false);

            // выход если нет колонок ФИО или Дата рождения 
            if (birthdayColumnIndex[0] == 0 || birthdayColumnIndex[1] == 0)
            {
                return;
            }

            // определяем количество записей по колонке ФИО
            int birthdayRowCount = MyRibbon.getRowCount(activeWorksheet, birthdayColumnIndex[0]);

            // получаем диапазон 
            Range birthdayWorksheetRange = birthdayWorksheet.Range[birthdayWorksheet.Cells[1, 1], birthdayWorksheet.Cells[birthdayRowCount, birthdayColumnIndex.Max()]];

            // перебираем все записи 
            DateTime dateBirthday;
            int posToInsert = 2;
            application.CutCopyMode = 0;
            for (int i = 2; i <= birthdayRowCount; i++)
            {
                if (DateTime.TryParse(birthdayWorksheetRange[i, birthdayColumnIndex[1]].Text, out dateBirthday))
                {
                    if (dateBirthday.Day >= date.Day)
                    {
                        birthdayWorksheet.Rows[posToInsert].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        birthdayWorksheet.Rows[posToInsert + 1].Copy();

                        goto skip1;
                    }
                    posToInsert++;
                }
            }

            birthdayWorksheet.Rows[posToInsert].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            birthdayWorksheet.Rows[posToInsert - 1].Copy();

        skip1:

            birthdayWorksheet.Rows[posToInsert].PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            application.CutCopyMode = 0;

            //birthdayWorksheet.Rows[posToInsert].Interior.Color = ColorTranslator.ToOle(Color.Red);

            for (int i = 0; i < birthdayColumnIndex.Length; i++)
            {
                if (birthdayColumnIndex[i] != 0 && activeWorksheetColumnIndex[i] != 0)
                {
                    birthdayWorksheet.Cells[posToInsert, birthdayColumnIndex[i]].value = activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[i]].value;
                }
            }

            StringBuilder sb = new StringBuilder()
                .Append("Запись успешно добавлена")
                .Append("\n\nФИО:\t")
                .Append(activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[0]].Text)
                .Append("\nДата:\t")
                .Append(activeWorksheet.Cells[activeCell.Row, activeWorksheetColumnIndex[1]].Text)
                .Append("\n\nЛист:\t")
                .Append(birthdayWorksheet.Name)
                .Append("\nСтрока:\t")
                .Append(posToInsert);

            MessageBox.Show(sb.ToString(), "Успех!", 0, MessageBoxIcon.Asterisk);
        }

        // добавить или отредактировать запись в интересах
        public void InterestsEntryEditor(Office.IRibbonControl control)
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.InterestsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            // получить индексы колонок по их заголовкам
            int[] headersIndex = MyRibbon.getHeadersIndex(activeWorksheet, visitsHeaders, 2);

            // если нет колонки ФИО 
            if (headersIndex[0] == 0) return;

            // получить количество строк в определенной колонке
            int rowCount = MyRibbon.getRowCount(activeWorksheet, headersIndex[0]);
            if (rowCount <= 2) return;

            // получаем диапазон 
            object[,] personNamesRange = activeWorksheet.Range[activeWorksheet.Cells[1, headersIndex[0]], activeWorksheet.Cells[rowCount, headersIndex[0]]].Value2;
            string[] personNames = new string[rowCount];
            for (int i = 1; i <= rowCount; i++)
            {
                if (personNamesRange[i, 1] != null)
                {
                    personNames[i - 1] = personNamesRange[i, 1].ToString();
                }
                else
                {
                    personNames[i - 1] = "";
                }
            }

            if (form_InterestsEntrySearch == null || form_InterestsEntrySearch.IsDisposed)
            {
                form_InterestsEntrySearch = new Form_InterestsEntrySearch(
                application,
                activeWorkBook,
                activeWorksheet,
                activeCell,
                headersIndex,
                personNames
            );
                form_InterestsEntrySearch.Show();
            }
            else
            {
                form_InterestsEntrySearch.WindowState = FormWindowState.Normal;
                form_InterestsEntrySearch.Activate();
            }
        }

        // Установки поведения надстройки
        public void SettingsEdit(Office.IRibbonControl control)
        {
            new Form_Settings().ShowDialog();

            WorkbookActivateHandler(null);
        }

        // -
        // Заполнить пропуски на всем листе
        public void GapsFill(Office.IRibbonControl control)
        {
            return;

            //Excel.Application application = Globals.ThisAddIn.Application;
            //Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            //string[] toReplace = new string[] { "нет данных", "не дала", "не дал", " ", "" };

            //Dictionary<string, List<string[]>> data = new Dictionary<string, List<string[]>>();

            //foreach (Excel.Workbook workbooks in application.Workbooks)
            //{
            //    if (workbooks.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) continue;

            //    foreach (Excel.Worksheet worksheet in workbooks.Worksheets)
            //    {
            //        // получить количество стобцов с заголовками 
            //        int columnCount = MyRibbon.getColumnCount(worksheet, 1);
            //        if (columnCount == 0) continue;

            //        // получить индексы колонок по их заголовкам
            //        int[] headersIndex = MyRibbon.getHeadersIndex(worksheet, visitsHeaders, 1);
            //        if (headersIndex[0] == 0) continue;

            //        // получить количество строк в определенной колонке
            //        int rowCount = MyRibbon.getRowCount(worksheet, headersIndex[0]);
            //        if (rowCount == 0) continue;

            //        if (rowCount == 1) rowCount++;

            //        // получаем диапазон 
            //        object[,] curRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, headersIndex.Max()]].value2;

            //        for (int j = 1; j <= rowCount; j++)
            //        {
            //            if (curRange[j, headersIndex[0]] != null)
            //            {
            //                string[] qwe = new string[headersIndex.Length];
            //                for (int i = 0; i < headersIndex.Length; i++)
            //                {
            //                    if (headersIndex[i] != 0 && curRange[j, headersIndex[i]] != null)
            //                    {
            //                        qwe[i] = curRange[j, headersIndex[i]].ToString();
            //                    }
            //                    else
            //                    {
            //                        qwe[i] = "";
            //                    }
            //                }

            //                string key = curRange[j, headersIndex[0]].ToString();
            //                if (data.ContainsKey(key))
            //                {
            //                    data[key].Add(qwe);
            //                }
            //                else
            //                {
            //                    data.Add(key, new List<string[]>());
            //                    data[key].Add(qwe);
            //                }
            //            }
            //        }
            //    }
            //}

            //// получить количество стобцов с заголовками 
            //int columnCount = MyRibbon.getColumnCount(activeWorksheet, 1);
            //if (columnCount == 0) return;

            //// получить индексы колонок по их заголовкам
            //int[] headersIndex = MyRibbon.getHeadersIndex(activeWorksheet, visitsHeaders, 1);
            //if (headersIndex[0] == 0) return;

            //// получить количество строк в определенной колонке
            //int rowCount = MyRibbon.getRowCount(activeWorksheet, headersIndex[0]);
            //if (rowCount == 0) return;
            //if (rowCount == 1) rowCount++;

            //object[,] curRange = activeWorksheet.Range[activeWorksheet.Cells[1, 1], activeWorksheet.Cells[rowCount, headersIndex.Max()]].value2;

            //for (int i = 2; i <= rowCount; i++)
            //{
            //    if (curRange[rowCount, headersIndex[0]] != null)
            //    {
            //        if (data.ContainsKey(curRange[rowCount, headersIndex[0]].ToString()))
            //        {
            //            for (int j = 1; j < headersIndex.Length; j++)
            //            {
            //                if (curRange[rowCount, headersIndex[j]] == null)
            //                {

            //                }
            //            }
            //        }

            //    }
            //}

            //return;
        }

        // Сделать поле «Участник...» везде как в первой записи ведомости посещений
        public void MakeIdentical(Office.IRibbonControl control)
        {
            DialogResult result = MessageBox.Show(
                "Сейчас будут получены данные из всех открытых книг \"Ведомость посещений\" и обработан весь активный лист.\n\nПродолжить?",
                "Внимание!",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result == DialogResult.No) return;

            Excel.Application application = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            if (activeWorkBook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) return;

            Dictionary<string, (DateTime, bool)> people = new Dictionary<string, (DateTime, bool)>();

            foreach (Excel.Workbook workbook in application.Workbooks)
            {
                if (workbook.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) continue;

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    // получить количество стобцов с заголовками 
                    int columnCount = MyRibbon.getColumnCount(worksheet, 1);
                    if (columnCount == 0) continue;

                    // получить индексы колонок по их заголовкам
                    int[] headersIndex = MyRibbon.getHeadersIndex(worksheet, makeIdenticalHeaders, 1);

                    // если нет колонок то выходим
                    if (headersIndex[0] == 0 || headersIndex[1] == 0 || headersIndex[2] == 0) continue;

                    // получить количество строк в определенной колонке
                    int rowCount = MyRibbon.getRowCount(worksheet, headersIndex[1]);
                    if (rowCount == 0) continue;

                    object[,] range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, headersIndex.Max()]].Value2;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        if (range[i, headersIndex[0]] is null || range[i, headersIndex[1]] is null || range[i, headersIndex[2]] is null) continue;

                        bool isParticipant;
                        if (string.Equals(range[i, headersIndex[2]].ToString(), "да", StringComparison.OrdinalIgnoreCase))
                        {
                            isParticipant = true;
                        }
                        else if (string.Equals(range[i, headersIndex[2]].ToString(), "нет", StringComparison.OrdinalIgnoreCase))
                        {
                            isParticipant = false;
                        }
                        else continue;

                        DateTime date;
                        if (range[i, headersIndex[0]] is double)
                        {
                            date = DateTime.FromOADate((double)range[i, headersIndex[0]]);
                        }
                        else
                        {
                            if (!DateTime.TryParse(range[i, headersIndex[0]].ToString(), out date))
                            {
                                continue;
                            }
                        }

                        string personName = range[i, headersIndex[1]].ToString();
                        if (people.ContainsKey(personName))
                        {
                            if (people[personName].Item2 != isParticipant)
                            {
                                if (people[personName].Item1 > date)
                                {
                                    people[personName] = (date, isParticipant);
                                }
                            }
                        }
                        else
                        {
                            people.Add(personName, (date, isParticipant));
                        }
                    }
                }
            }

            // получить количество стобцов с заголовками 
            int activeColumnCount = MyRibbon.getColumnCount(activeWorksheet, 1);
            if (activeColumnCount == 0) return;

            // получить индексы колонок по их заголовкам
            int[] activeHadersIndex = MyRibbon.getHeadersIndex(activeWorksheet, makeIdenticalHeaders, 1);

            // если нет колонок то выходим
            if (activeHadersIndex[0] == 0 || activeHadersIndex[1] == 0 || activeHadersIndex[2] == 0) return;

            // получить количество строк в определенной колонке
            int activeRowCount = MyRibbon.getRowCount(activeWorksheet, activeHadersIndex[1]);

            if (activeRowCount == 0) return;

            object[,] activeRange = activeWorksheet.Range[activeWorksheet.Cells[1, 1], activeWorksheet.Cells[activeRowCount, activeHadersIndex.Max()]].Value2;

            for (int i = 1; i <= activeRowCount; i++)
            {
                if (activeRange[i, activeHadersIndex[1]] is null) continue;

                string personName = activeRange[i, activeHadersIndex[1]].ToString();

                if (people.ContainsKey(personName))
                {
                    if (people[personName].Item2)
                    {
                        activeWorksheet.Cells[i, activeHadersIndex[2]].value = "Да";
                    }
                    else
                    {
                        activeWorksheet.Cells[i, activeHadersIndex[2]].value = "Нет";
                    }
                    //activeWorksheet.Cells[i, headersIndex[2]].Interior.Color = ColorTranslator.ToOle(Color.Red);
                }
            }

            MessageBox.Show("Обработка завершена", "Успех!", 0, MessageBoxIcon.Asterisk);
        }

        // Удалить лишние пробелы
        public void RemoveExtraSpaces(Office.IRibbonControl control)
        {
            DialogResult result = MessageBox.Show(
                "Сейчас будет произведен поиск во всех ячейках листа с текстом и удалены лишние пробелы\n(более 2х пробелов между словами и пробелы перед и после текста)\n\nПродолжить?",
                "Внимание!",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result == DialogResult.Yes)
            {
                try
                {
                    Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

                    // получить количество стобцов с заголовками 
                    int columnCount = 0;
                    for (int i = 1; i <= 10; i++)
                    {
                        columnCount = Math.Max(columnCount, MyRibbon.getColumnCount(activeWorksheet, i));
                    }
                    if (columnCount == 0) return;

                    // получить количество строк 
                    int rowCount = 0;
                    for (int i = 1; i <= columnCount; i++)
                    {
                        rowCount = Math.Max(rowCount, MyRibbon.getRowCount(activeWorksheet, i));
                    }
                    if (rowCount == 0) return;

                    if (columnCount == 1 && rowCount == 1) rowCount++;

                    object[,] table = activeWorksheet.Range[activeWorksheet.Cells[1, 1], activeWorksheet.Cells[rowCount, columnCount]].value2;

                    int countOfDeleted = 0;
                    double progress = 1;
                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= columnCount; j++)
                        {
                            if (table[i, j] is string)
                            {
                                string qwe = System.Text.RegularExpressions.Regex.Replace(table[i, j].ToString(), @"[ ]+", " ").Trim();

                                if (table[i, j].ToString() != qwe)
                                {
                                    activeWorksheet.Cells[i, j].value = qwe;
                                    countOfDeleted++;
                                }
                            }
                        }
                    }

                    MessageBox.Show($"Изменено записей: {countOfDeleted}", "Успех!", 0, MessageBoxIcon.Asterisk);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Ошибка!", 0, MessageBoxIcon.Error);
                }
            }
        }

        // Удалить переносы строк в ячейках
        public void RemoveNewLine(Office.IRibbonControl control)
        {
            DialogResult result = MessageBox.Show(
                "Сейчас будет произведен поиск во всех ячейках листа с текстом и удалены все переносы строк в одной ячейке\n\nПродолжить?",
                "Внимание!",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result == DialogResult.Yes)
            {
                try
                {
                    Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

                    // получить количество стобцов с заголовками 
                    int columnCount = MyRibbon.getColumnCount(activeWorksheet, 1);
                    if (columnCount == 0) return;

                    // получить количество строк 
                    int rowCount = 0;
                    for (int i = 1; i <= columnCount; i++)
                    {
                        rowCount = Math.Max(rowCount, MyRibbon.getRowCount(activeWorksheet, i));
                    }
                    if (rowCount == 0) return;

                    if (columnCount == 1 && rowCount == 1) rowCount++;

                    object[,] table = activeWorksheet.Range[activeWorksheet.Cells[1, 1], activeWorksheet.Cells[rowCount, columnCount]].value2;

                    int countOfDeleted = 0;
                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= columnCount; j++)
                        {
                            if (table[i, j] is string)
                            {
                                string qwe = System.Text.RegularExpressions.Regex.Replace(table[i, j].ToString(), @"[\n]+", "");

                                if (table[i, j].ToString() != qwe)
                                {
                                    activeWorksheet.Cells[i, j].value = qwe;
                                    countOfDeleted++;
                                }
                            }
                        }
                    }

                    MessageBox.Show($"Изменено записей: {countOfDeleted}", "Успех!", 0, MessageBoxIcon.Asterisk);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Ошибка!", 0, MessageBoxIcon.Error);
                }
            }
        }

        // -
        // Исправление опечаток в выделеном столбце
        public void TyposCorrection(Office.IRibbonControl control)
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;

            //new Form_TyposCorrection(application, activeWorkBook, activeWorksheet, activeCell).ShowDialog();

            return;
        }

        #endregion

        #region Вспомогательные методы

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        //
        public System.Drawing.Image GetImage(string ImageName)
        {
            return (System.Drawing.Image)Properties.Resources.ResourceManager.GetObject(ImageName);
        }

        // получить количество стобцов с заголовками 
        public static int getColumnCount(Excel.Worksheet worksheet, int headersRow)
        {
            int columnCount = 1;
            int countOfEmpty = 0;

            object[,] worksheetRow = worksheet.Rows[headersRow].value2;

            while (true)
            {
                if (worksheetRow[1, columnCount] is null)
                {
                    countOfEmpty++;
                }
                else
                {
                    countOfEmpty = 0;
                }

                if (countOfEmpty == 25)
                {
                    break;
                }
                columnCount++;
            }

            columnCount -= 25;

            return columnCount;
        }

        // получить индексы колонок по их заголовкам
        public static int[] getHeadersIndex(Excel.Worksheet worksheet, string[] headers, int headersRow, bool exactMatch = true)
        {
            int headersLength = headers.Length;

            int[] headersIndex = new int[headersLength];

            object[,] worksheetRow = worksheet.Rows[headersRow].value2;

            for (int i = 1; i <= worksheetRow.Length; i++)
            {
                if (worksheetRow[1, i] is null) continue;

                for (int j = 0; j < headersLength; j++)
                {
                    if (exactMatch)
                    {
                        if (worksheetRow[1, i].ToString() == headers[j])
                        {
                            headersIndex[j] = i;
                            break;
                        }
                    }
                    else
                    {
                        if (worksheetRow[1, i].ToString().IndexOf(headers[j], StringComparison.OrdinalIgnoreCase) >= 0 ||
                            headers[j].IndexOf(worksheetRow[1, i].ToString(), StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            headersIndex[j] = i;
                            break;
                        }
                    }
                }
            }
            return headersIndex;
        }

        // получить количество строк в определенной колонке
        public static int getRowCount(Excel.Worksheet worksheet, int columnIndex)
        {
            int rowCount = 1;
            int rowIndex = 1;
            int countOfEmpty = 0;

            object[,] worksheetColumn = worksheet.Range[worksheet.Cells[1, columnIndex], worksheet.Cells[500, columnIndex]].value2;

            while (true)
            {
                if (worksheetColumn[rowIndex, 1] is null)
                {
                    countOfEmpty++;
                }
                else
                {
                    countOfEmpty = 0;
                }

                if (countOfEmpty == 25)
                {
                    break;
                }

                if (rowIndex == 500)
                {
                    rowIndex = 0;
                    worksheetColumn = worksheet.Range[worksheet.Cells[rowCount + 1, columnIndex], worksheet.Cells[rowCount + 500, columnIndex]].value2;
                }

                rowIndex++;
                rowCount++;
            }

            rowCount -= 25;

            return rowCount;
        }

        // 
        public static string[,] GetDataToStringArray(Excel.Application application, bool withoutActiveRow = false)
        {
            int totalRowCount = 0;

            Dictionary<string, (int, int[])> listMetadata = new Dictionary<string, (int, int[])>();

            foreach (Excel.Workbook workbooks in application.Workbooks)
            {
                if (workbooks.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) continue;

                foreach (Excel.Worksheet worksheet in workbooks.Worksheets)
                {
                    // получить количество стобцов с заголовками 
                    int columnCount = MyRibbon.getColumnCount(worksheet, 1);
                    if (columnCount == 0) continue;

                    // получить индексы колонок по их заголовкам
                    int[] headersIndex = MyRibbon.getHeadersIndex(worksheet, visitsHeaders, 1);
                    if (headersIndex[0] == 0) continue;

                    // получить количество строк в определенной колонке
                    int rowCount = MyRibbon.getRowCount(worksheet, headersIndex[0]);
                    if (rowCount <= 1) continue;

                    listMetadata.Add(workbooks.Name + worksheet.Name, (rowCount, headersIndex));

                    totalRowCount += rowCount;
                }
            }

            string[,] qwe = new string[totalRowCount, visitsHeaders.Length + 1];
            int curPos = 0;

            string activeWorkbookName = application.ActiveWorkbook.Name;
            string activeSheetName = application.ActiveWorkbook.ActiveSheet.Name;
            int activeRow = application.ActiveCell.Row;

            foreach (Excel.Workbook workbooks in application.Workbooks)
            {
                if (workbooks.Name.IndexOf(Properties.Settings.Default.VisitsFileName, StringComparison.OrdinalIgnoreCase) < 0) continue;

                foreach (Excel.Worksheet worksheet in workbooks.Worksheets)
                {
                    string workbooksName = workbooks.Name;
                    string key = workbooksName + worksheet.Name;

                    bool onActiveSheet = (activeWorkbookName == workbooks.Name && activeSheetName == worksheet.Name);

                    if (listMetadata.ContainsKey(key))
                    {
                        int rowCount = listMetadata[key].Item1;
                        int[] headersIndex = listMetadata[key].Item2;

                        // получаем диапазон 
                        object[,] curRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, headersIndex.Max()]].value2;

                        for (int j = 1; j <= rowCount; j++)
                        {
                            if (withoutActiveRow && onActiveSheet && j == activeRow) continue;

                            for (int i = 0; i < headersIndex.Length; i++)
                            {

                                if (headersIndex[i] != 0 && curRange[j, headersIndex[i]] != null)
                                {
                                    if (i == 2 && curRange[j, headersIndex[i]] is double)
                                    {
                                        try
                                        {
                                            qwe[curPos + j - 1, i] = DateTime.FromOADate((double)curRange[j, headersIndex[i]]).ToString("dd.MM.yyyy");
                                        }
                                        catch (Exception)
                                        {
                                            qwe[curPos + j - 1, i] = curRange[j, headersIndex[i]].ToString();
                                        }
                                    }
                                    else
                                    {
                                        qwe[curPos + j - 1, i] = curRange[j, headersIndex[i]].ToString();
                                    }
                                }
                                else
                                {
                                    qwe[curPos + j - 1, i] = "";
                                }

                            }
                            qwe[curPos + j - 1, headersIndex.Length] = workbooksName;
                        }
                        curPos += rowCount - 1;
                    }
                }
            }

            return qwe;
        }

        // 
        public static List<List<string[]>> findEntryInStringArray(string[,] stringArray, string textToFinding, bool insertSurname, bool withoutDuplicating)
        {
            List<List<string[]>> collectionOfElements = new List<List<string[]>>();
            for (int i = 0; i < 8; i++) collectionOfElements.Add(new List<string[]>());

            int curItem;
            string elem;
            string[] qwe;

            for (int i = 1; i < stringArray.GetLength(0); i++)
            {
                elem = stringArray[i, 0];

                if (string.IsNullOrEmpty(elem)) continue;

                curItem = -1;
                //точное совпадение
                if (string.Equals(elem, textToFinding, StringComparison.OrdinalIgnoreCase)) curItem = 0;
                //точное совпадение с фамилией 
                else if (string.Equals(elem.Split(' ')[0], textToFinding, StringComparison.OrdinalIgnoreCase)) curItem = 1;
                //точное совпадение фамилии с фамилией
                else if (insertSurname && string.Equals(elem.Split(' ')[0], textToFinding.Split(' ')[0], StringComparison.OrdinalIgnoreCase)) curItem = 2;
                //точное совпадение с началом фамилии
                else if (elem.Length >= textToFinding.Length && string.Equals(elem.Substring(0, textToFinding.Length), textToFinding, StringComparison.OrdinalIgnoreCase)) curItem = 3;
                //вхождение в фамилию       
                else if (elem.Split(' ')[0].IndexOf(textToFinding, StringComparison.OrdinalIgnoreCase) >= 0) curItem = 4;
                //вхождение фамилии в фамилию
                else if (insertSurname && elem.Split(' ')[0].IndexOf(textToFinding.Split(' ')[0], StringComparison.OrdinalIgnoreCase) >= 0) curItem = 5;
                //вхождение в ФИО
                else if (elem.IndexOf(textToFinding, StringComparison.OrdinalIgnoreCase) >= 0) curItem = 6;
                //вхождение фамилии в ФИО
                else if (insertSurname && elem.IndexOf(textToFinding.Split(' ')[0], StringComparison.OrdinalIgnoreCase) >= 0) curItem = 7;

                if (curItem >= 0)
                {
                    qwe = new string[visitsHeaders.Length + 1];

                    for (int j = 0; j < visitsHeaders.Length + 1; j++)
                    {
                        qwe[j] = stringArray[i, j];
                    }

                    if (withoutDuplicating)
                    {
                        if (!isInCollection(collectionOfElements[curItem], qwe, true))
                        {
                            collectionOfElements[curItem].Add(qwe);
                        }
                    }
                    else
                    {
                        collectionOfElements[curItem].Add(qwe);
                    }
                }
            }

            return collectionOfElements;
        }

        //
        public static bool isInCollection(List<string[]> collection, string[] elem, bool withoutLast)
        {
            int len = elem.Length;

            if (withoutLast) len--;

            bool flag;
            foreach (string[] element in collection)
            {
                if (element.Length == elem.Length)
                {
                    flag = true;
                    for (int i = 0; i < len; i++)
                    {
                        if ((element[i] ?? string.Empty).Trim() != (elem[i] ?? string.Empty).Trim())
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        // 
        public static string[] getHeadersArray(List<(string, int)> visitsHeadersAndWeight, bool withoutLast)
        {
            int len = visitsHeadersAndWeight.Count;

            if (withoutLast) len--;

            string[] visitsHeaders = new string[len];

            for (int i = 0; i < len; i++)
            {
                visitsHeaders[i] = visitsHeadersAndWeight[i].Item1;
            }

            return visitsHeaders;
        }

        #endregion
    }
}
