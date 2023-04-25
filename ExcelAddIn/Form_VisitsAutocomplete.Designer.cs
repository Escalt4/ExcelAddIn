namespace ExcelAddIn
{
    partial class Form_VisitsAutocomplete
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button_addSelected = new System.Windows.Forms.Button();
            this.checkBox_findSurname = new System.Windows.Forms.CheckBox();
            this.listView = new System.Windows.Forms.ListView();
            this.checkBox_withoutDuplicatingEntrys = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button_addSelected
            // 
            this.button_addSelected.Enabled = false;
            this.button_addSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_addSelected.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button_addSelected.Location = new System.Drawing.Point(749, 377);
            this.button_addSelected.Name = "button_addSelected";
            this.button_addSelected.Size = new System.Drawing.Size(223, 72);
            this.button_addSelected.TabIndex = 1;
            this.button_addSelected.Text = "Взять выбранную запись";
            this.button_addSelected.UseVisualStyleBackColor = true;
            this.button_addSelected.Click += new System.EventHandler(this.button_addSelected_Click);
            // 
            // checkBox_findSurname
            // 
            this.checkBox_findSurname.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBox_findSurname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_findSurname.Location = new System.Drawing.Point(15, 360);
            this.checkBox_findSurname.Name = "checkBox_findSurname";
            this.checkBox_findSurname.Size = new System.Drawing.Size(213, 40);
            this.checkBox_findSurname.TabIndex = 3;
            this.checkBox_findSurname.Text = "Искать в том числе по фамилии";
            this.checkBox_findSurname.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBox_findSurname.UseVisualStyleBackColor = true;
            this.checkBox_findSurname.CheckedChanged += new System.EventHandler(this.checkBox_findSurname_CheckedChanged);
            // 
            // listView
            // 
            this.listView.FullRowSelect = true;
            this.listView.GridLines = true;
            this.listView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView.HideSelection = false;
            this.listView.Location = new System.Drawing.Point(12, 12);
            this.listView.MultiSelect = false;
            this.listView.Name = "listView";
            this.listView.ShowGroups = false;
            this.listView.Size = new System.Drawing.Size(960, 342);
            this.listView.TabIndex = 6;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            this.listView.SelectedIndexChanged += new System.EventHandler(this.listView_SelectedIndexChanged);
            // 
            // checkBox_withoutDuplicatingEntrys
            // 
            this.checkBox_withoutDuplicatingEntrys.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBox_withoutDuplicatingEntrys.Checked = true;
            this.checkBox_withoutDuplicatingEntrys.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_withoutDuplicatingEntrys.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_withoutDuplicatingEntrys.Location = new System.Drawing.Point(15, 417);
            this.checkBox_withoutDuplicatingEntrys.Name = "checkBox_withoutDuplicatingEntrys";
            this.checkBox_withoutDuplicatingEntrys.Size = new System.Drawing.Size(213, 32);
            this.checkBox_withoutDuplicatingEntrys.TabIndex = 8;
            this.checkBox_withoutDuplicatingEntrys.Text = "Исключить одинаковые записи";
            this.checkBox_withoutDuplicatingEntrys.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBox_withoutDuplicatingEntrys.UseVisualStyleBackColor = true;
            this.checkBox_withoutDuplicatingEntrys.CheckedChanged += new System.EventHandler(this.checkBox_withoutDuplicatingEntrys_CheckedChanged);
            // 
            // Form_visitsAutocomplete
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 461);
            this.Controls.Add(this.checkBox_withoutDuplicatingEntrys);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.checkBox_findSurname);
            this.Controls.Add(this.button_addSelected);
            this.MaximizeBox = false;
            this.Name = "Form_visitsAutocomplete";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ручное автозаполнение";
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button_addSelected;
        private System.Windows.Forms.CheckBox checkBox_findSurname;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.CheckBox checkBox_withoutDuplicatingEntrys;
    }
}