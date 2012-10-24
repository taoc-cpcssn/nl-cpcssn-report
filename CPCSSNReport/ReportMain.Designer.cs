namespace CPCSSNReport
{
    partial class ReportMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportMain));
            this.cmdByProvider = new System.Windows.Forms.Button();
            this.cmdAll = new System.Windows.Forms.Button();
            this.cmdByPractice = new System.Windows.Forms.Button();
            this.cboProviders = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboPractices = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDBPrev = new System.Windows.Forms.TextBox();
            this.txtDBCurrent = new System.Windows.Forms.TextBox();
            this.cmdConnectDB = new System.Windows.Forms.Button();
            this.txtTemplate = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cmdAppendAllSummary = new System.Windows.Forms.Button();
            this.cmdSaveExcel = new System.Windows.Forms.Button();
            this.cmdAppendByProvider = new System.Windows.Forms.Button();
            this.cmdAppendByPractice = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.txtExcelTemplate = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmdByProvider
            // 
            this.cmdByProvider.Location = new System.Drawing.Point(187, 438);
            this.cmdByProvider.Name = "cmdByProvider";
            this.cmdByProvider.Size = new System.Drawing.Size(133, 26);
            this.cmdByProvider.TabIndex = 0;
            this.cmdByProvider.Text = "By Provider";
            this.cmdByProvider.UseVisualStyleBackColor = true;
            this.cmdByProvider.Click += new System.EventHandler(this.cmdByProvider_Click);
            // 
            // cmdAll
            // 
            this.cmdAll.Location = new System.Drawing.Point(21, 438);
            this.cmdAll.Name = "cmdAll";
            this.cmdAll.Size = new System.Drawing.Size(133, 26);
            this.cmdAll.TabIndex = 1;
            this.cmdAll.Text = "All Summary";
            this.cmdAll.UseVisualStyleBackColor = true;
            this.cmdAll.Click += new System.EventHandler(this.cmdAll_Click);
            // 
            // cmdByPractice
            // 
            this.cmdByPractice.Location = new System.Drawing.Point(358, 438);
            this.cmdByPractice.Name = "cmdByPractice";
            this.cmdByPractice.Size = new System.Drawing.Size(133, 26);
            this.cmdByPractice.TabIndex = 2;
            this.cmdByPractice.Text = "By Practice";
            this.cmdByPractice.UseVisualStyleBackColor = true;
            this.cmdByPractice.Click += new System.EventHandler(this.cmdByPractice_Click);
            // 
            // cboProviders
            // 
            this.cboProviders.FormattingEnabled = true;
            this.cboProviders.Location = new System.Drawing.Point(187, 282);
            this.cboProviders.Name = "cboProviders";
            this.cboProviders.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.cboProviders.Size = new System.Drawing.Size(129, 147);
            this.cboProviders.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(187, 266);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Provider List";
            // 
            // cboPractices
            // 
            this.cboPractices.FormattingEnabled = true;
            this.cboPractices.Location = new System.Drawing.Point(358, 281);
            this.cboPractices.Name = "cboPractices";
            this.cboPractices.Size = new System.Drawing.Size(133, 21);
            this.cboPractices.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(358, 266);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Practice List";
            // 
            // txtDBPrev
            // 
            this.txtDBPrev.Location = new System.Drawing.Point(21, 70);
            this.txtDBPrev.Name = "txtDBPrev";
            this.txtDBPrev.Size = new System.Drawing.Size(337, 20);
            this.txtDBPrev.TabIndex = 8;
            this.txtDBPrev.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtDBPrev_MouseDown);
            // 
            // txtDBCurrent
            // 
            this.txtDBCurrent.Location = new System.Drawing.Point(21, 24);
            this.txtDBCurrent.Name = "txtDBCurrent";
            this.txtDBCurrent.Size = new System.Drawing.Size(337, 20);
            this.txtDBCurrent.TabIndex = 7;
            this.txtDBCurrent.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtDBCurrent_MouseDown);
            // 
            // cmdConnectDB
            // 
            this.cmdConnectDB.Location = new System.Drawing.Point(396, 20);
            this.cmdConnectDB.Name = "cmdConnectDB";
            this.cmdConnectDB.Size = new System.Drawing.Size(106, 26);
            this.cmdConnectDB.TabIndex = 11;
            this.cmdConnectDB.Text = "Connect DB";
            this.cmdConnectDB.UseVisualStyleBackColor = true;
            this.cmdConnectDB.Click += new System.EventHandler(this.cmdConnectDB_Click);
            // 
            // txtTemplate
            // 
            this.txtTemplate.Location = new System.Drawing.Point(21, 118);
            this.txtTemplate.Name = "txtTemplate";
            this.txtTemplate.Size = new System.Drawing.Size(336, 20);
            this.txtTemplate.TabIndex = 12;
            this.txtTemplate.MouseDown += new System.Windows.Forms.MouseEventHandler(this.textBox1_MouseDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 8);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Current DB";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "Previous DB";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(21, 100);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Template";
            // 
            // cmdAppendAllSummary
            // 
            this.cmdAppendAllSummary.Location = new System.Drawing.Point(21, 478);
            this.cmdAppendAllSummary.Name = "cmdAppendAllSummary";
            this.cmdAppendAllSummary.Size = new System.Drawing.Size(133, 25);
            this.cmdAppendAllSummary.TabIndex = 17;
            this.cmdAppendAllSummary.Text = "Append All Summary";
            this.cmdAppendAllSummary.UseVisualStyleBackColor = true;
            this.cmdAppendAllSummary.Click += new System.EventHandler(this.cmdAppend_Click);
            // 
            // cmdSaveExcel
            // 
            this.cmdSaveExcel.Location = new System.Drawing.Point(358, 517);
            this.cmdSaveExcel.Name = "cmdSaveExcel";
            this.cmdSaveExcel.Size = new System.Drawing.Size(133, 25);
            this.cmdSaveExcel.TabIndex = 18;
            this.cmdSaveExcel.Text = "Save Excel";
            this.cmdSaveExcel.UseVisualStyleBackColor = true;
            this.cmdSaveExcel.Click += new System.EventHandler(this.cmdSaveExcel_Click);
            // 
            // cmdAppendByProvider
            // 
            this.cmdAppendByProvider.Location = new System.Drawing.Point(187, 478);
            this.cmdAppendByProvider.Name = "cmdAppendByProvider";
            this.cmdAppendByProvider.Size = new System.Drawing.Size(133, 25);
            this.cmdAppendByProvider.TabIndex = 19;
            this.cmdAppendByProvider.Text = "Append By Provider";
            this.cmdAppendByProvider.UseVisualStyleBackColor = true;
            this.cmdAppendByProvider.Click += new System.EventHandler(this.cmdAppendByProvider_Click);
            // 
            // cmdAppendByPractice
            // 
            this.cmdAppendByPractice.Location = new System.Drawing.Point(358, 478);
            this.cmdAppendByPractice.Name = "cmdAppendByPractice";
            this.cmdAppendByPractice.Size = new System.Drawing.Size(133, 25);
            this.cmdAppendByPractice.TabIndex = 20;
            this.cmdAppendByPractice.Text = "Append By Practice";
            this.cmdAppendByPractice.UseVisualStyleBackColor = true;
            this.cmdAppendByPractice.Click += new System.EventHandler(this.cmdAppendByPractice_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(21, 155);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 13);
            this.label6.TabIndex = 21;
            this.label6.Text = "Excel Template";
            // 
            // txtExcelTemplate
            // 
            this.txtExcelTemplate.Location = new System.Drawing.Point(21, 173);
            this.txtExcelTemplate.Name = "txtExcelTemplate";
            this.txtExcelTemplate.Size = new System.Drawing.Size(336, 20);
            this.txtExcelTemplate.TabIndex = 22;
            this.txtExcelTemplate.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtExcelTemplate_MouseDown);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(21, 211);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 13);
            this.label7.TabIndex = 23;
            this.label7.Text = "Output Path";
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(22, 228);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(336, 20);
            this.txtOutput.TabIndex = 24;
            this.txtOutput.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtOutput_MouseDown);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(17, 517);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(303, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "Multile selection in Provider List produces a report for the group";
            // 
            // ReportMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(503, 555);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtExcelTemplate);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cmdAppendByPractice);
            this.Controls.Add(this.cmdAppendByProvider);
            this.Controls.Add(this.cmdSaveExcel);
            this.Controls.Add(this.cmdAppendAllSummary);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtTemplate);
            this.Controls.Add(this.cmdConnectDB);
            this.Controls.Add(this.txtDBPrev);
            this.Controls.Add(this.txtDBCurrent);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cboPractices);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboProviders);
            this.Controls.Add(this.cmdByPractice);
            this.Controls.Add(this.cmdAll);
            this.Controls.Add(this.cmdByProvider);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ReportMain";
            this.Text = "CPCSSN Auto Report";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cmdByProvider;
        private System.Windows.Forms.Button cmdAll;
        private System.Windows.Forms.Button cmdByPractice;
        private System.Windows.Forms.ListBox cboProviders;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboPractices;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDBPrev;
        private System.Windows.Forms.TextBox txtDBCurrent;
        private System.Windows.Forms.Button cmdConnectDB;
        private System.Windows.Forms.TextBox txtTemplate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button cmdAppendAllSummary;
        private System.Windows.Forms.Button cmdSaveExcel;
        private System.Windows.Forms.Button cmdAppendByProvider;
        private System.Windows.Forms.Button cmdAppendByPractice;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtExcelTemplate;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Label label8;
    }
}

