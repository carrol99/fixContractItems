namespace FixContractItems
{
    partial class frmFixContractItems
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.mnuExit = new System.Windows.Forms.ToolStripMenuItem();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabMismatched = new System.Windows.Forms.TabPage();
            this.splitMismatchedMain = new System.Windows.Forms.SplitContainer();
            this.btnWriteMoveSQL = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnWriteDeleteSQL = new System.Windows.Forms.Button();
            this.lblFoundCount = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblNotFoundCount = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnValidate = new System.Windows.Forms.Button();
            this.btnRetrieveMismatched = new System.Windows.Forms.Button();
            this.lblMismatchedRecordCount = new System.Windows.Forms.Label();
            this.btnWriteMismatched = new System.Windows.Forms.Button();
            this.lblMismatchSheetNameDesc = new System.Windows.Forms.Label();
            this.lblMismatchedOutputFileDesc = new System.Windows.Forms.Label();
            this.btnMismatchedExpand = new System.Windows.Forms.Button();
            this.dgvMismatched = new System.Windows.Forms.DataGridView();
            this.tabMissingItems = new System.Windows.Forms.TabPage();
            this.splitMissingMain = new System.Windows.Forms.SplitContainer();
            this.label5 = new System.Windows.Forms.Label();
            this.txtOutputFileName = new System.Windows.Forms.TextBox();
            this.txtMismatchedSheetName = new System.Windows.Forms.TextBox();
            this.txtMismatchedOutputFileName = new System.Windows.Forms.TextBox();
            this.txtMissingFileName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtMissingSheetName = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.tabMain.SuspendLayout();
            this.tabMismatched.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitMismatchedMain)).BeginInit();
            this.splitMismatchedMain.Panel1.SuspendLayout();
            this.splitMismatchedMain.Panel2.SuspendLayout();
            this.splitMismatchedMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMismatched)).BeginInit();
            this.tabMissingItems.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitMissingMain)).BeginInit();
            this.splitMissingMain.Panel1.SuspendLayout();
            this.splitMissingMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuExit});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1244, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // mnuExit
            // 
            this.mnuExit.Name = "mnuExit";
            this.mnuExit.Size = new System.Drawing.Size(38, 20);
            this.mnuExit.Text = "Exit";
            this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
            // 
            // tabMain
            // 
            this.tabMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabMain.Controls.Add(this.tabMismatched);
            this.tabMain.Controls.Add(this.tabMissingItems);
            this.tabMain.Location = new System.Drawing.Point(0, 27);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(1244, 597);
            this.tabMain.TabIndex = 1;
            // 
            // tabMismatched
            // 
            this.tabMismatched.Controls.Add(this.splitMismatchedMain);
            this.tabMismatched.Location = new System.Drawing.Point(4, 22);
            this.tabMismatched.Name = "tabMismatched";
            this.tabMismatched.Padding = new System.Windows.Forms.Padding(3);
            this.tabMismatched.Size = new System.Drawing.Size(1236, 571);
            this.tabMismatched.TabIndex = 0;
            this.tabMismatched.Text = "Mismatched Items";
            this.tabMismatched.UseVisualStyleBackColor = true;
            // 
            // splitMismatchedMain
            // 
            this.splitMismatchedMain.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitMismatchedMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitMismatchedMain.IsSplitterFixed = true;
            this.splitMismatchedMain.Location = new System.Drawing.Point(3, 3);
            this.splitMismatchedMain.Name = "splitMismatchedMain";
            this.splitMismatchedMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitMismatchedMain.Panel1
            // 
            this.splitMismatchedMain.Panel1.Controls.Add(this.btnWriteMoveSQL);
            this.splitMismatchedMain.Panel1.Controls.Add(this.txtOutputFileName);
            this.splitMismatchedMain.Panel1.Controls.Add(this.label4);
            this.splitMismatchedMain.Panel1.Controls.Add(this.btnWriteDeleteSQL);
            this.splitMismatchedMain.Panel1.Controls.Add(this.lblFoundCount);
            this.splitMismatchedMain.Panel1.Controls.Add(this.label3);
            this.splitMismatchedMain.Panel1.Controls.Add(this.lblNotFoundCount);
            this.splitMismatchedMain.Panel1.Controls.Add(this.label2);
            this.splitMismatchedMain.Panel1.Controls.Add(this.label1);
            this.splitMismatchedMain.Panel1.Controls.Add(this.lblStatus);
            this.splitMismatchedMain.Panel1.Controls.Add(this.btnValidate);
            this.splitMismatchedMain.Panel1.Controls.Add(this.btnRetrieveMismatched);
            // 
            // splitMismatchedMain.Panel2
            // 
            this.splitMismatchedMain.Panel2.Controls.Add(this.lblMismatchedRecordCount);
            this.splitMismatchedMain.Panel2.Controls.Add(this.btnWriteMismatched);
            this.splitMismatchedMain.Panel2.Controls.Add(this.txtMismatchedSheetName);
            this.splitMismatchedMain.Panel2.Controls.Add(this.txtMismatchedOutputFileName);
            this.splitMismatchedMain.Panel2.Controls.Add(this.lblMismatchSheetNameDesc);
            this.splitMismatchedMain.Panel2.Controls.Add(this.lblMismatchedOutputFileDesc);
            this.splitMismatchedMain.Panel2.Controls.Add(this.btnMismatchedExpand);
            this.splitMismatchedMain.Panel2.Controls.Add(this.dgvMismatched);
            this.splitMismatchedMain.Size = new System.Drawing.Size(1230, 565);
            this.splitMismatchedMain.SplitterDistance = 143;
            this.splitMismatchedMain.TabIndex = 0;
            // 
            // btnWriteMoveSQL
            // 
            this.btnWriteMoveSQL.Location = new System.Drawing.Point(494, 18);
            this.btnWriteMoveSQL.Name = "btnWriteMoveSQL";
            this.btnWriteMoveSQL.Size = new System.Drawing.Size(103, 23);
            this.btnWriteMoveSQL.TabIndex = 11;
            this.btnWriteMoveSQL.Text = "Write Move SQL";
            this.btnWriteMoveSQL.UseVisualStyleBackColor = true;
            this.btnWriteMoveSQL.Click += new System.EventHandler(this.btnWriteMoveSQL_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 60);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Output File Name";
            // 
            // btnWriteDeleteSQL
            // 
            this.btnWriteDeleteSQL.Location = new System.Drawing.Point(365, 19);
            this.btnWriteDeleteSQL.Name = "btnWriteDeleteSQL";
            this.btnWriteDeleteSQL.Size = new System.Drawing.Size(107, 23);
            this.btnWriteDeleteSQL.TabIndex = 8;
            this.btnWriteDeleteSQL.Text = "Write Delete SQL";
            this.btnWriteDeleteSQL.UseVisualStyleBackColor = true;
            this.btnWriteDeleteSQL.Click += new System.EventHandler(this.btnWriteDeleteSQL_Click);
            // 
            // lblFoundCount
            // 
            this.lblFoundCount.AutoSize = true;
            this.lblFoundCount.Location = new System.Drawing.Point(897, 22);
            this.lblFoundCount.Name = "lblFoundCount";
            this.lblFoundCount.Size = new System.Drawing.Size(13, 13);
            this.lblFoundCount.TabIndex = 7;
            this.lblFoundCount.Text = "0";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(850, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Found:";
            // 
            // lblNotFoundCount
            // 
            this.lblNotFoundCount.AutoSize = true;
            this.lblNotFoundCount.Location = new System.Drawing.Point(805, 22);
            this.lblNotFoundCount.Name = "lblNotFoundCount";
            this.lblNotFoundCount.Size = new System.Drawing.Size(13, 13);
            this.lblNotFoundCount.TabIndex = 5;
            this.lblNotFoundCount.Text = "0";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(741, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Not found:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(613, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Read:";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(676, 23);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(13, 13);
            this.lblStatus.TabIndex = 2;
            this.lblStatus.Text = "0";
            // 
            // btnValidate
            // 
            this.btnValidate.Location = new System.Drawing.Point(230, 18);
            this.btnValidate.Name = "btnValidate";
            this.btnValidate.Size = new System.Drawing.Size(93, 23);
            this.btnValidate.TabIndex = 1;
            this.btnValidate.Text = "Validate Items";
            this.btnValidate.UseVisualStyleBackColor = true;
            this.btnValidate.Click += new System.EventHandler(this.btnValidate_Click);
            // 
            // btnRetrieveMismatched
            // 
            this.btnRetrieveMismatched.Location = new System.Drawing.Point(8, 19);
            this.btnRetrieveMismatched.Name = "btnRetrieveMismatched";
            this.btnRetrieveMismatched.Size = new System.Drawing.Size(162, 23);
            this.btnRetrieveMismatched.TabIndex = 0;
            this.btnRetrieveMismatched.Text = "Retrieve Mismatched Items";
            this.btnRetrieveMismatched.UseVisualStyleBackColor = true;
            this.btnRetrieveMismatched.Click += new System.EventHandler(this.btnRetrieveMismatched_Click);
            // 
            // lblMismatchedRecordCount
            // 
            this.lblMismatchedRecordCount.AutoSize = true;
            this.lblMismatchedRecordCount.Location = new System.Drawing.Point(701, 56);
            this.lblMismatchedRecordCount.Name = "lblMismatchedRecordCount";
            this.lblMismatchedRecordCount.Size = new System.Drawing.Size(13, 13);
            this.lblMismatchedRecordCount.TabIndex = 7;
            this.lblMismatchedRecordCount.Text = "0";
            // 
            // btnWriteMismatched
            // 
            this.btnWriteMismatched.Location = new System.Drawing.Point(442, 47);
            this.btnWriteMismatched.Name = "btnWriteMismatched";
            this.btnWriteMismatched.Size = new System.Drawing.Size(97, 23);
            this.btnWriteMismatched.TabIndex = 6;
            this.btnWriteMismatched.Text = "Write Excel File";
            this.btnWriteMismatched.UseVisualStyleBackColor = true;
            this.btnWriteMismatched.Click += new System.EventHandler(this.btnWriteMismatched_Click);
            // 
            // lblMismatchSheetNameDesc
            // 
            this.lblMismatchSheetNameDesc.AutoSize = true;
            this.lblMismatchSheetNameDesc.Location = new System.Drawing.Point(18, 47);
            this.lblMismatchSheetNameDesc.Name = "lblMismatchSheetNameDesc";
            this.lblMismatchSheetNameDesc.Size = new System.Drawing.Size(66, 13);
            this.lblMismatchSheetNameDesc.TabIndex = 3;
            this.lblMismatchSheetNameDesc.Text = "Sheet Name";
            // 
            // lblMismatchedOutputFileDesc
            // 
            this.lblMismatchedOutputFileDesc.AutoSize = true;
            this.lblMismatchedOutputFileDesc.Location = new System.Drawing.Point(15, 16);
            this.lblMismatchedOutputFileDesc.Name = "lblMismatchedOutputFileDesc";
            this.lblMismatchedOutputFileDesc.Size = new System.Drawing.Size(118, 13);
            this.lblMismatchedOutputFileDesc.TabIndex = 2;
            this.lblMismatchedOutputFileDesc.Text = "Output Excel File Name";
            // 
            // btnMismatchedExpand
            // 
            this.btnMismatchedExpand.Location = new System.Drawing.Point(622, 47);
            this.btnMismatchedExpand.Name = "btnMismatchedExpand";
            this.btnMismatchedExpand.Size = new System.Drawing.Size(27, 26);
            this.btnMismatchedExpand.TabIndex = 1;
            this.btnMismatchedExpand.Text = " +";
            this.btnMismatchedExpand.UseVisualStyleBackColor = true;
            this.btnMismatchedExpand.Click += new System.EventHandler(this.btnMismatchedExpand_Click);
            // 
            // dgvMismatched
            // 
            this.dgvMismatched.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvMismatched.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMismatched.Location = new System.Drawing.Point(3, 94);
            this.dgvMismatched.Name = "dgvMismatched";
            this.dgvMismatched.Size = new System.Drawing.Size(1220, 317);
            this.dgvMismatched.TabIndex = 0;
            // 
            // tabMissingItems
            // 
            this.tabMissingItems.Controls.Add(this.splitMissingMain);
            this.tabMissingItems.Location = new System.Drawing.Point(4, 22);
            this.tabMissingItems.Name = "tabMissingItems";
            this.tabMissingItems.Padding = new System.Windows.Forms.Padding(3);
            this.tabMissingItems.Size = new System.Drawing.Size(1236, 571);
            this.tabMissingItems.TabIndex = 1;
            this.tabMissingItems.Text = "Missing Items";
            this.tabMissingItems.UseVisualStyleBackColor = true;
            // 
            // splitMissingMain
            // 
            this.splitMissingMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitMissingMain.Location = new System.Drawing.Point(3, 6);
            this.splitMissingMain.Name = "splitMissingMain";
            this.splitMissingMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitMissingMain.Panel1
            // 
            this.splitMissingMain.Panel1.Controls.Add(this.txtMissingSheetName);
            this.splitMissingMain.Panel1.Controls.Add(this.label6);
            this.splitMissingMain.Panel1.Controls.Add(this.txtMissingFileName);
            this.splitMissingMain.Panel1.Controls.Add(this.label5);
            this.splitMissingMain.Size = new System.Drawing.Size(1230, 562);
            this.splitMissingMain.SplitterDistance = 155;
            this.splitMissingMain.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 15);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Premium File";
            // 
            // txtOutputFileName
            // 
            this.txtOutputFileName.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::FixContractItems.Properties.Settings.Default, "txtOutputFileName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtOutputFileName.Location = new System.Drawing.Point(113, 53);
            this.txtOutputFileName.Name = "txtOutputFileName";
            this.txtOutputFileName.Size = new System.Drawing.Size(1100, 20);
            this.txtOutputFileName.TabIndex = 10;
            this.txtOutputFileName.Text = global::FixContractItems.Properties.Settings.Default.txtOutputFileName;
            // 
            // txtMismatchedSheetName
            // 
            this.txtMismatchedSheetName.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::FixContractItems.Properties.Settings.Default, "txtMismatchedSheetName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtMismatchedSheetName.Location = new System.Drawing.Point(139, 47);
            this.txtMismatchedSheetName.Name = "txtMismatchedSheetName";
            this.txtMismatchedSheetName.Size = new System.Drawing.Size(234, 20);
            this.txtMismatchedSheetName.TabIndex = 5;
            this.txtMismatchedSheetName.Text = global::FixContractItems.Properties.Settings.Default.txtMismatchedSheetName;
            // 
            // txtMismatchedOutputFileName
            // 
            this.txtMismatchedOutputFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtMismatchedOutputFileName.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::FixContractItems.Properties.Settings.Default, "txtMismatchedOutputFileName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtMismatchedOutputFileName.Location = new System.Drawing.Point(139, 13);
            this.txtMismatchedOutputFileName.Name = "txtMismatchedOutputFileName";
            this.txtMismatchedOutputFileName.Size = new System.Drawing.Size(1058, 20);
            this.txtMismatchedOutputFileName.TabIndex = 4;
            this.txtMismatchedOutputFileName.Text = global::FixContractItems.Properties.Settings.Default.txtMismatchedOutputFileName;
            // 
            // txtMissingFileName
            // 
            this.txtMissingFileName.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::FixContractItems.Properties.Settings.Default, "txtMissingFileName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtMissingFileName.Location = new System.Drawing.Point(18, 42);
            this.txtMissingFileName.Name = "txtMissingFileName";
            this.txtMissingFileName.Size = new System.Drawing.Size(1192, 20);
            this.txtMissingFileName.TabIndex = 1;
            this.txtMissingFileName.Text = global::FixContractItems.Properties.Settings.Default.txtMissingFileName;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 72);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(66, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Sheet Name";
            // 
            // txtMissingSheetName
            // 
            this.txtMissingSheetName.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::FixContractItems.Properties.Settings.Default, "txtMissingSheetName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtMissingSheetName.Location = new System.Drawing.Point(91, 69);
            this.txtMissingSheetName.Name = "txtMissingSheetName";
            this.txtMissingSheetName.Size = new System.Drawing.Size(212, 20);
            this.txtMissingSheetName.TabIndex = 3;
            this.txtMissingSheetName.Text = global::FixContractItems.Properties.Settings.Default.txtMissingSheetName;
            // 
            // frmFixContractItems
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1244, 626);
            this.Controls.Add(this.tabMain);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmFixContractItems";
            this.Text = "frmFixContractItems";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabMain.ResumeLayout(false);
            this.tabMismatched.ResumeLayout(false);
            this.splitMismatchedMain.Panel1.ResumeLayout(false);
            this.splitMismatchedMain.Panel1.PerformLayout();
            this.splitMismatchedMain.Panel2.ResumeLayout(false);
            this.splitMismatchedMain.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitMismatchedMain)).EndInit();
            this.splitMismatchedMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMismatched)).EndInit();
            this.tabMissingItems.ResumeLayout(false);
            this.splitMissingMain.Panel1.ResumeLayout(false);
            this.splitMissingMain.Panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitMissingMain)).EndInit();
            this.splitMissingMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem mnuExit;
        private System.Windows.Forms.TabControl tabMain;
        private System.Windows.Forms.TabPage tabMismatched;
        private System.Windows.Forms.SplitContainer splitMismatchedMain;
        private System.Windows.Forms.Button btnMismatchedExpand;
        private System.Windows.Forms.DataGridView dgvMismatched;
        private System.Windows.Forms.TabPage tabMissingItems;
        private System.Windows.Forms.TextBox txtMismatchedOutputFileName;
        private System.Windows.Forms.Label lblMismatchSheetNameDesc;
        private System.Windows.Forms.Label lblMismatchedOutputFileDesc;
        private System.Windows.Forms.Button btnRetrieveMismatched;
        private System.Windows.Forms.Button btnWriteMismatched;
        private System.Windows.Forms.TextBox txtMismatchedSheetName;
        private System.Windows.Forms.Label lblMismatchedRecordCount;
        private System.Windows.Forms.Button btnValidate;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblFoundCount;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblNotFoundCount;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnWriteDeleteSQL;
        private System.Windows.Forms.TextBox txtOutputFileName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnWriteMoveSQL;
        private System.Windows.Forms.SplitContainer splitMissingMain;
        private System.Windows.Forms.TextBox txtMissingFileName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtMissingSheetName;
        private System.Windows.Forms.Label label6;
    }
}

