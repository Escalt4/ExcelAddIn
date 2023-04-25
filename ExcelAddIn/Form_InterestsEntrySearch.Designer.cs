namespace ExcelAddIn
{
    partial class Form_InterestsEntrySearch
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
            this.listView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.textBox = new System.Windows.Forms.TextBox();
            this.button_EditChanged = new System.Windows.Forms.Button();
            this.checkBox_findSurname = new System.Windows.Forms.CheckBox();
            this.button_CreateNew = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listView
            // 
            this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listView.FullRowSelect = true;
            this.listView.GridLines = true;
            this.listView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listView.HideSelection = false;
            this.listView.Location = new System.Drawing.Point(21, 76);
            this.listView.MultiSelect = false;
            this.listView.Name = "listView";
            this.listView.ShowGroups = false;
            this.listView.Size = new System.Drawing.Size(543, 581);
            this.listView.TabIndex = 15;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            this.listView.SelectedIndexChanged += new System.EventHandler(this.listView_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Width = 515;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(166, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Поиск по ФИО (от 3 символов)";
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(21, 37);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(543, 20);
            this.textBox.TabIndex = 10;
            this.textBox.TextChanged += new System.EventHandler(this.textBox_TextChanged);
            // 
            // button_EditChanged
            // 
            this.button_EditChanged.Enabled = false;
            this.button_EditChanged.Location = new System.Drawing.Point(21, 703);
            this.button_EditChanged.Name = "button_EditChanged";
            this.button_EditChanged.Size = new System.Drawing.Size(255, 46);
            this.button_EditChanged.TabIndex = 18;
            this.button_EditChanged.Text = "Редактировать выбранную";
            this.button_EditChanged.UseVisualStyleBackColor = true;
            this.button_EditChanged.Click += new System.EventHandler(this.button_EditChanged_Click);
            // 
            // checkBox_findSurname
            // 
            this.checkBox_findSurname.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox_findSurname.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.checkBox_findSurname.Location = new System.Drawing.Point(28, 663);
            this.checkBox_findSurname.Name = "checkBox_findSurname";
            this.checkBox_findSurname.Size = new System.Drawing.Size(213, 34);
            this.checkBox_findSurname.TabIndex = 17;
            this.checkBox_findSurname.Text = "Искать в том числе по фамилии";
            this.checkBox_findSurname.UseVisualStyleBackColor = true;
            this.checkBox_findSurname.CheckedChanged += new System.EventHandler(this.checkBox_findSurname_CheckedChanged);
            // 
            // button_CreateNew
            // 
            this.button_CreateNew.Location = new System.Drawing.Point(309, 703);
            this.button_CreateNew.Name = "button_CreateNew";
            this.button_CreateNew.Size = new System.Drawing.Size(255, 46);
            this.button_CreateNew.TabIndex = 16;
            this.button_CreateNew.Text = "Создать новую";
            this.button_CreateNew.UseVisualStyleBackColor = true;
            this.button_CreateNew.Click += new System.EventHandler(this.button_CreateNew_Click);
            // 
            // Form_InterestsEntrySearch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 761);
            this.Controls.Add(this.button_EditChanged);
            this.Controls.Add(this.checkBox_findSurname);
            this.Controls.Add(this.button_CreateNew);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox);
            this.MaximizeBox = false;
            this.Name = "Form_InterestsEntrySearch";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Поиск записи в таблице \"Интересы посетителей\"";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.Button button_EditChanged;
        private System.Windows.Forms.CheckBox checkBox_findSurname;
        private System.Windows.Forms.Button button_CreateNew;
    }
}