namespace PCMExceltoSQLDatabase
{
    partial class PCMExceltoSQLDatabase
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PCMExceltoSQLDatabase));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label4 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.Progress = new System.Windows.Forms.Label();
            this.EndRowlabel = new System.Windows.Forms.Label();
            this.StartingRowlabel = new System.Windows.Forms.Label();
            this.EndRow = new System.Windows.Forms.TextBox();
            this.StartingRow = new System.Windows.Forms.TextBox();
            this.Update = new System.Windows.Forms.CheckBox();
            this.Insert = new System.Windows.Forms.CheckBox();
            this.Insertsubmit = new System.Windows.Forms.Button();
            this.Assetscheck = new System.Windows.Forms.CheckBox();
            this.Activitycheck = new System.Windows.Forms.CheckBox();
            this.Peformancecheck = new System.Windows.Forms.CheckBox();
            this.Flowscheck = new System.Windows.Forms.CheckBox();
            this.Accountscheck = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.Status = new System.Windows.Forms.Label();
            this.Passwordlabel = new System.Windows.Forms.Label();
            this.Usernamelabel = new System.Windows.Forms.Label();
            this.Password = new System.Windows.Forms.TextBox();
            this.Username = new System.Windows.Forms.TextBox();
            this.ExcelWorksheetName = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.ExcelWorksheetNameLabel = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.firstpagesubmit = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.printtext = new System.Windows.Forms.Label();
            this.InsertUpdateTable = new System.Windows.Forms.TabControl();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.InsertUpdateTable.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(612, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 7;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.Progress);
            this.tabPage2.Controls.Add(this.EndRowlabel);
            this.tabPage2.Controls.Add(this.StartingRowlabel);
            this.tabPage2.Controls.Add(this.EndRow);
            this.tabPage2.Controls.Add(this.StartingRow);
            this.tabPage2.Controls.Add(this.Update);
            this.tabPage2.Controls.Add(this.Insert);
            this.tabPage2.Controls.Add(this.Insertsubmit);
            this.tabPage2.Controls.Add(this.Assetscheck);
            this.tabPage2.Controls.Add(this.Activitycheck);
            this.tabPage2.Controls.Add(this.Peformancecheck);
            this.tabPage2.Controls.Add(this.Flowscheck);
            this.tabPage2.Controls.Add(this.Accountscheck);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(675, 561);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Insert or Update Table";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // Progress
            // 
            this.Progress.AutoSize = true;
            this.Progress.Location = new System.Drawing.Point(273, 480);
            this.Progress.Name = "Progress";
            this.Progress.Size = new System.Drawing.Size(0, 13);
            this.Progress.TabIndex = 13;
            // 
            // EndRowlabel
            // 
            this.EndRowlabel.AutoSize = true;
            this.EndRowlabel.Location = new System.Drawing.Point(406, 169);
            this.EndRowlabel.Name = "EndRowlabel";
            this.EndRowlabel.Size = new System.Drawing.Size(74, 13);
            this.EndRowlabel.TabIndex = 12;
            this.EndRowlabel.Text = "Row to end at";
            // 
            // StartingRowlabel
            // 
            this.StartingRowlabel.AutoSize = true;
            this.StartingRowlabel.Location = new System.Drawing.Point(403, 75);
            this.StartingRowlabel.Name = "StartingRowlabel";
            this.StartingRowlabel.Size = new System.Drawing.Size(76, 13);
            this.StartingRowlabel.TabIndex = 11;
            this.StartingRowlabel.Text = "Row to start at";
            // 
            // EndRow
            // 
            this.EndRow.Location = new System.Drawing.Point(403, 206);
            this.EndRow.Name = "EndRow";
            this.EndRow.Size = new System.Drawing.Size(100, 20);
            this.EndRow.TabIndex = 10;
            // 
            // StartingRow
            // 
            this.StartingRow.Location = new System.Drawing.Point(403, 110);
            this.StartingRow.Name = "StartingRow";
            this.StartingRow.Size = new System.Drawing.Size(100, 20);
            this.StartingRow.TabIndex = 9;
            // 
            // Update
            // 
            this.Update.AutoSize = true;
            this.Update.Location = new System.Drawing.Point(342, 363);
            this.Update.Name = "Update";
            this.Update.Size = new System.Drawing.Size(61, 17);
            this.Update.TabIndex = 8;
            this.Update.Text = "Update";
            this.Update.UseVisualStyleBackColor = true;
            // 
            // Insert
            // 
            this.Insert.AutoSize = true;
            this.Insert.Location = new System.Drawing.Point(221, 363);
            this.Insert.Name = "Insert";
            this.Insert.Size = new System.Drawing.Size(52, 17);
            this.Insert.TabIndex = 7;
            this.Insert.Text = "Insert";
            this.Insert.UseVisualStyleBackColor = true;
            // 
            // Insertsubmit
            // 
            this.Insertsubmit.Location = new System.Drawing.Point(273, 414);
            this.Insertsubmit.Name = "Insertsubmit";
            this.Insertsubmit.Size = new System.Drawing.Size(75, 23);
            this.Insertsubmit.TabIndex = 6;
            this.Insertsubmit.Text = "Submit";
            this.Insertsubmit.UseVisualStyleBackColor = true;
            this.Insertsubmit.Click += new System.EventHandler(this.Insertsubmit_Click);
            // 
            // Assetscheck
            // 
            this.Assetscheck.AutoSize = true;
            this.Assetscheck.Location = new System.Drawing.Point(103, 296);
            this.Assetscheck.Name = "Assetscheck";
            this.Assetscheck.Size = new System.Drawing.Size(57, 17);
            this.Assetscheck.TabIndex = 5;
            this.Assetscheck.Text = "Assets";
            this.Assetscheck.UseVisualStyleBackColor = true;
            // 
            // Activitycheck
            // 
            this.Activitycheck.AutoSize = true;
            this.Activitycheck.Location = new System.Drawing.Point(103, 251);
            this.Activitycheck.Name = "Activitycheck";
            this.Activitycheck.Size = new System.Drawing.Size(60, 17);
            this.Activitycheck.TabIndex = 4;
            this.Activitycheck.Text = "Activity";
            this.Activitycheck.UseVisualStyleBackColor = true;
            // 
            // Peformancecheck
            // 
            this.Peformancecheck.AutoSize = true;
            this.Peformancecheck.Location = new System.Drawing.Point(103, 206);
            this.Peformancecheck.Name = "Peformancecheck";
            this.Peformancecheck.Size = new System.Drawing.Size(86, 17);
            this.Peformancecheck.TabIndex = 3;
            this.Peformancecheck.Text = "Performance";
            this.Peformancecheck.UseVisualStyleBackColor = true;
            // 
            // Flowscheck
            // 
            this.Flowscheck.AutoSize = true;
            this.Flowscheck.Location = new System.Drawing.Point(103, 157);
            this.Flowscheck.Name = "Flowscheck";
            this.Flowscheck.Size = new System.Drawing.Size(53, 17);
            this.Flowscheck.TabIndex = 2;
            this.Flowscheck.Text = "Flows";
            this.Flowscheck.UseVisualStyleBackColor = true;
            // 
            // Accountscheck
            // 
            this.Accountscheck.AutoSize = true;
            this.Accountscheck.Location = new System.Drawing.Point(103, 110);
            this.Accountscheck.Name = "Accountscheck";
            this.Accountscheck.Size = new System.Drawing.Size(71, 17);
            this.Accountscheck.TabIndex = 1;
            this.Accountscheck.Text = "Accounts";
            this.Accountscheck.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(166, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(300, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Complete Database and Excel File Section before proceeding!";
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.Status);
            this.tabPage1.Controls.Add(this.Passwordlabel);
            this.tabPage1.Controls.Add(this.Usernamelabel);
            this.tabPage1.Controls.Add(this.Password);
            this.tabPage1.Controls.Add(this.Username);
            this.tabPage1.Controls.Add(this.ExcelWorksheetName);
            this.tabPage1.Controls.Add(this.textBox2);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.ExcelWorksheetNameLabel);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.firstpagesubmit);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.printtext);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(675, 561);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Database and Excel File Selection";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // Status
            // 
            this.Status.AutoSize = true;
            this.Status.Location = new System.Drawing.Point(159, 447);
            this.Status.Name = "Status";
            this.Status.Size = new System.Drawing.Size(0, 13);
            this.Status.TabIndex = 13;
            // 
            // Passwordlabel
            // 
            this.Passwordlabel.AutoSize = true;
            this.Passwordlabel.Location = new System.Drawing.Point(30, 308);
            this.Passwordlabel.Name = "Passwordlabel";
            this.Passwordlabel.Size = new System.Drawing.Size(53, 13);
            this.Passwordlabel.TabIndex = 12;
            this.Passwordlabel.Text = "Password";
            // 
            // Usernamelabel
            // 
            this.Usernamelabel.AutoSize = true;
            this.Usernamelabel.Location = new System.Drawing.Point(30, 256);
            this.Usernamelabel.Name = "Usernamelabel";
            this.Usernamelabel.Size = new System.Drawing.Size(55, 13);
            this.Usernamelabel.TabIndex = 11;
            this.Usernamelabel.Text = "Username";
            // 
            // Password
            // 
            this.Password.Location = new System.Drawing.Point(230, 302);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.Size = new System.Drawing.Size(100, 20);
            this.Password.TabIndex = 10;
            // 
            // Username
            // 
            this.Username.Location = new System.Drawing.Point(230, 250);
            this.Username.Name = "Username";
            this.Username.Size = new System.Drawing.Size(100, 20);
            this.Username.TabIndex = 9;
            // 
            // ExcelWorksheetName
            // 
            this.ExcelWorksheetName.Location = new System.Drawing.Point(230, 85);
            this.ExcelWorksheetName.Name = "ExcelWorksheetName";
            this.ExcelWorksheetName.Size = new System.Drawing.Size(100, 20);
            this.ExcelWorksheetName.TabIndex = 8;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(230, 199);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 20);
            this.textBox2.TabIndex = 6;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(230, 132);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 20);
            this.textBox1.TabIndex = 1;
            // 
            // ExcelWorksheetNameLabel
            // 
            this.ExcelWorksheetNameLabel.AutoSize = true;
            this.ExcelWorksheetNameLabel.Location = new System.Drawing.Point(30, 85);
            this.ExcelWorksheetNameLabel.Name = "ExcelWorksheetNameLabel";
            this.ExcelWorksheetNameLabel.Size = new System.Drawing.Size(147, 13);
            this.ExcelWorksheetNameLabel.TabIndex = 7;
            this.ExcelWorksheetNameLabel.Text = "Enter Excel Worksheet Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select the Excel File";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(230, 40);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Select File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // firstpagesubmit
            // 
            this.firstpagesubmit.Location = new System.Drawing.Point(147, 379);
            this.firstpagesubmit.Name = "firstpagesubmit";
            this.firstpagesubmit.Size = new System.Drawing.Size(75, 23);
            this.firstpagesubmit.TabIndex = 2;
            this.firstpagesubmit.Text = "Submit";
            this.firstpagesubmit.UseVisualStyleBackColor = true;
            this.firstpagesubmit.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 135);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Enter the Database name ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 199);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Enter the table name ";
            // 
            // printtext
            // 
            this.printtext.AutoSize = true;
            this.printtext.Location = new System.Drawing.Point(75, 420);
            this.printtext.Name = "printtext";
            this.printtext.Size = new System.Drawing.Size(0, 13);
            this.printtext.TabIndex = 4;
            // 
            // InsertUpdateTable
            // 
            this.InsertUpdateTable.Controls.Add(this.tabPage1);
            this.InsertUpdateTable.Controls.Add(this.tabPage2);
            this.InsertUpdateTable.Location = new System.Drawing.Point(-1, 0);
            this.InsertUpdateTable.Name = "InsertUpdateTable";
            this.InsertUpdateTable.SelectedIndex = 0;
            this.InsertUpdateTable.Size = new System.Drawing.Size(683, 587);
            this.InsertUpdateTable.TabIndex = 8;
            this.InsertUpdateTable.Click += new System.EventHandler(this.OpenInsertUpdateTable);
            // 
            // PCMExceltoSQLDatabase
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(676, 582);
            this.Controls.Add(this.InsertUpdateTable);
            this.Controls.Add(this.label4);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PCMExceltoSQLDatabase";
            this.Text = "Polen Capital Management Excel to SQL Database Application";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.InsertUpdateTable.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label EndRowlabel;
        private System.Windows.Forms.Label StartingRowlabel;
        private System.Windows.Forms.TextBox EndRow;
        private System.Windows.Forms.TextBox StartingRow;
        private System.Windows.Forms.CheckBox Update;
        private System.Windows.Forms.CheckBox Insert;
        private System.Windows.Forms.Button Insertsubmit;
        private System.Windows.Forms.CheckBox Assetscheck;
        private System.Windows.Forms.CheckBox Activitycheck;
        private System.Windows.Forms.CheckBox Peformancecheck;
        private System.Windows.Forms.CheckBox Flowscheck;
        private System.Windows.Forms.CheckBox Accountscheck;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox ExcelWorksheetName;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label ExcelWorksheetNameLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button firstpagesubmit;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label printtext;
        private System.Windows.Forms.TabControl InsertUpdateTable;
        private System.Windows.Forms.Label Passwordlabel;
        private System.Windows.Forms.Label Usernamelabel;
        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.TextBox Username;
        private System.Windows.Forms.Label Status;
        private System.Windows.Forms.Label Progress;
    }
}

