namespace ExcelAddIn
{
    partial class Form_VisitsAddEntry
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
            this.label = new System.Windows.Forms.Label();
            this.button_addSelected = new System.Windows.Forms.Button();
            this.button_addNew = new System.Windows.Forms.Button();
            this.checkBox_addCurTime = new System.Windows.Forms.CheckBox();
            this.checkBox_toEndOfList = new System.Windows.Forms.CheckBox();
            this.listView = new System.Windows.Forms.ListView();
            this.textBox = new System.Windows.Forms.TextBox();
            this.checkBox_withoutDuplicatingEntrys = new System.Windows.Forms.CheckBox();
            this.checkBox_findSurname = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Location = new System.Drawing.Point(18, 21);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(166, 13);
            this.label.TabIndex = 0;
            this.label.Text = "Поиск по ФИО (от 3 символов)";
            // 
            // button_addSelected
            // 
            this.button_addSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_addSelected.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button_addSelected.Location = new System.Drawing.Point(749, 402);
            this.button_addSelected.Name = "button_addSelected";
            this.button_addSelected.Size = new System.Drawing.Size(223, 72);
            this.button_addSelected.TabIndex = 1;
            this.button_addSelected.Text = "Взять выбранную запись";
            this.button_addSelected.UseVisualStyleBackColor = true;
            this.button_addSelected.Click += new System.EventHandler(this.button_addSelected_Click);
            // 
            // button_addNew
            // 
            this.button_addNew.Location = new System.Drawing.Point(749, 37);
            this.button_addNew.Name = "button_addNew";
            this.button_addNew.Size = new System.Drawing.Size(223, 42);
            this.button_addNew.TabIndex = 2;
            this.button_addNew.Text = "Добавить как новую";
            this.button_addNew.UseVisualStyleBackColor = true;
            this.button_addNew.Click += new System.EventHandler(this.button_addNew_Click);
            // 
            // checkBox_addCurTime
            // 
            this.checkBox_addCurTime.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBox_addCurTime.Checked = true;
            this.checkBox_addCurTime.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_addCurTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_addCurTime.Location = new System.Drawing.Point(596, 402);
            this.checkBox_addCurTime.Name = "checkBox_addCurTime";
            this.checkBox_addCurTime.Size = new System.Drawing.Size(150, 27);
            this.checkBox_addCurTime.TabIndex = 4;
            this.checkBox_addCurTime.Text = "Указать текущее время";
            this.checkBox_addCurTime.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBox_addCurTime.UseVisualStyleBackColor = true;
            // 
            // checkBox_toEndOfList
            // 
            this.checkBox_toEndOfList.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBox_toEndOfList.Checked = true;
            this.checkBox_toEndOfList.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_toEndOfList.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_toEndOfList.Location = new System.Drawing.Point(596, 446);
            this.checkBox_toEndOfList.Name = "checkBox_toEndOfList";
            this.checkBox_toEndOfList.Size = new System.Drawing.Size(147, 28);
            this.checkBox_toEndOfList.TabIndex = 5;
            this.checkBox_toEndOfList.Text = "В конец списка";
            this.checkBox_toEndOfList.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBox_toEndOfList.UseVisualStyleBackColor = true;
            this.checkBox_toEndOfList.CheckedChanged += new System.EventHandler(this.checkBox_toEndOfList_CheckedChanged);
            // 
            // listView
            // 
            this.listView.FullRowSelect = true;
            this.listView.GridLines = true;
            this.listView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView.HideSelection = false;
            this.listView.Location = new System.Drawing.Point(14, 94);
            this.listView.MultiSelect = false;
            this.listView.Name = "listView";
            this.listView.ShowGroups = false;
            this.listView.Size = new System.Drawing.Size(958, 285);
            this.listView.TabIndex = 6;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            this.listView.SelectedIndexChanged += new System.EventHandler(this.listView_SelectedIndexChanged);
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(14, 48);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(715, 20);
            this.textBox.TabIndex = 7;
            this.textBox.TextChanged += new System.EventHandler(this.textBox_TextChanged);
            // 
            // checkBox_withoutDuplicatingEntrys
            // 
            this.checkBox_withoutDuplicatingEntrys.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBox_withoutDuplicatingEntrys.Checked = true;
            this.checkBox_withoutDuplicatingEntrys.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_withoutDuplicatingEntrys.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_withoutDuplicatingEntrys.Location = new System.Drawing.Point(21, 446);
            this.checkBox_withoutDuplicatingEntrys.Name = "checkBox_withoutDuplicatingEntrys";
            this.checkBox_withoutDuplicatingEntrys.Size = new System.Drawing.Size(213, 32);
            this.checkBox_withoutDuplicatingEntrys.TabIndex = 10;
            this.checkBox_withoutDuplicatingEntrys.Text = "Исключить одинаковые записи";
            this.checkBox_withoutDuplicatingEntrys.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBox_withoutDuplicatingEntrys.UseVisualStyleBackColor = true;
            this.checkBox_withoutDuplicatingEntrys.CheckedChanged += new System.EventHandler(this.checkBox_withoutDuplicatingEntrys_CheckedChanged);
            // 
            // checkBox_findSurname
            // 
            this.checkBox_findSurname.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBox_findSurname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_findSurname.Location = new System.Drawing.Point(21, 389);
            this.checkBox_findSurname.Name = "checkBox_findSurname";
            this.checkBox_findSurname.Size = new System.Drawing.Size(213, 40);
            this.checkBox_findSurname.TabIndex = 9;
            this.checkBox_findSurname.Text = "Искать в том числе по фамилии";
            this.checkBox_findSurname.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBox_findSurname.UseVisualStyleBackColor = true;
            this.checkBox_findSurname.CheckedChanged += new System.EventHandler(this.checkBox_findSurname_CheckedChanged);
            // 
            // Form_visitsAddEntry
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 486);
            this.Controls.Add(this.checkBox_withoutDuplicatingEntrys);
            this.Controls.Add(this.checkBox_findSurname);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.checkBox_toEndOfList);
            this.Controls.Add(this.checkBox_addCurTime);
            this.Controls.Add(this.button_addNew);
            this.Controls.Add(this.button_addSelected);
            this.Controls.Add(this.label);
            this.MaximizeBox = false;
            this.Name = "Form_visitsAddEntry";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавление новой записи";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label;
        private System.Windows.Forms.Button button_addSelected;
        private System.Windows.Forms.Button button_addNew;
        private System.Windows.Forms.CheckBox checkBox_addCurTime;
        private System.Windows.Forms.CheckBox checkBox_toEndOfList;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.CheckBox checkBox_withoutDuplicatingEntrys;
        private System.Windows.Forms.CheckBox checkBox_findSurname;
    }
}