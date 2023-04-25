using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class Form_Settings : Form
    {
        public Form_Settings()
        {
            InitializeComponent();

            textBox_VisitsFileName.Text = Properties.Settings.Default.VisitsFileName;
            textBox_BirthdayFileName.Text = Properties.Settings.Default.BirthdayFileName;
            textBox_InterestsFileName.Text = Properties.Settings.Default.InterestsFileName;
        }

        private void button_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.VisitsFileName = textBox_VisitsFileName.Text;
            Properties.Settings.Default.BirthdayFileName = textBox_BirthdayFileName.Text;
            Properties.Settings.Default.InterestsFileName = textBox_InterestsFileName.Text;
            Properties.Settings.Default.Save();

            Close();
        }
    }
}
