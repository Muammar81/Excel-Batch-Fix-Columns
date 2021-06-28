
namespace OBCFix
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.chkOpenAfterFix = new System.Windows.Forms.CheckBox();
            this.fixButton = new System.Windows.Forms.Button();
            this.sheetsList = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pwdBox = new System.Windows.Forms.TextBox();
            this.passwordCheckBox = new System.Windows.Forms.CheckBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.chkFixDates = new System.Windows.Forms.CheckBox();
            this.chkIgnoreCols = new System.Windows.Forms.CheckBox();
            this.txtDateHeader = new System.Windows.Forms.TextBox();
            this.txtIgnoredCols = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkCleanHeaderNames = new System.Windows.Forms.CheckBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.lblColumns = new System.Windows.Forms.Label();
            this.numHeaderOffset = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.lvCols = new System.Windows.Forms.ListView();
            this.contextMenueIgnoreDate = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ignoreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numHeaderOffset)).BeginInit();
            this.contextMenueIgnoreDate.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.button1.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.25F);
            this.button1.Location = new System.Drawing.Point(12, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(200, 48);
            this.button1.TabIndex = 1;
            this.button1.Text = "Browse for Excel Files";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.RestoreDirectory = true;
            // 
            // chkOpenAfterFix
            // 
            this.chkOpenAfterFix.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.chkOpenAfterFix.AutoSize = true;
            this.chkOpenAfterFix.Checked = true;
            this.chkOpenAfterFix.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOpenAfterFix.Font = new System.Drawing.Font("Arial Narrow", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOpenAfterFix.ForeColor = System.Drawing.Color.White;
            this.chkOpenAfterFix.Location = new System.Drawing.Point(12, 315);
            this.chkOpenAfterFix.Name = "chkOpenAfterFix";
            this.chkOpenAfterFix.Size = new System.Drawing.Size(159, 19);
            this.chkOpenAfterFix.TabIndex = 2;
            this.chkOpenAfterFix.Text = "Open Output Folder After Fix";
            this.chkOpenAfterFix.UseVisualStyleBackColor = true;
            // 
            // fixButton
            // 
            this.fixButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.fixButton.Enabled = false;
            this.fixButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.fixButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.fixButton.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.25F);
            this.fixButton.Location = new System.Drawing.Point(12, 279);
            this.fixButton.Name = "fixButton";
            this.fixButton.Size = new System.Drawing.Size(485, 23);
            this.fixButton.TabIndex = 1;
            this.fixButton.Text = "No Files Selected...";
            this.fixButton.UseVisualStyleBackColor = true;
            this.fixButton.Click += new System.EventHandler(this.fixButton_Click);
            // 
            // sheetsList
            // 
            this.sheetsList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sheetsList.FormattingEnabled = true;
            this.sheetsList.Location = new System.Drawing.Point(13, 178);
            this.sheetsList.Name = "sheetsList";
            this.sheetsList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.sheetsList.Size = new System.Drawing.Size(199, 95);
            this.sheetsList.TabIndex = 4;
            this.sheetsList.SelectedIndexChanged += new System.EventHandler(this.sheetsList_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 157);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Sheets";
            // 
            // pwdBox
            // 
            this.pwdBox.Enabled = false;
            this.pwdBox.Location = new System.Drawing.Point(12, 103);
            this.pwdBox.Name = "pwdBox";
            this.pwdBox.Size = new System.Drawing.Size(163, 20);
            this.pwdBox.TabIndex = 6;
            this.pwdBox.UseSystemPasswordChar = true;
            // 
            // passwordCheckBox
            // 
            this.passwordCheckBox.AutoSize = true;
            this.passwordCheckBox.Location = new System.Drawing.Point(12, 80);
            this.passwordCheckBox.Name = "passwordCheckBox";
            this.passwordCheckBox.Size = new System.Drawing.Size(175, 17);
            this.passwordCheckBox.TabIndex = 7;
            this.passwordCheckBox.Text = "Sheets are Password Protected";
            this.passwordCheckBox.UseVisualStyleBackColor = true;
            this.passwordCheckBox.CheckedChanged += new System.EventHandler(this.passwordCheckBox_CheckedChanged);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::OBCFix.Properties.Resources.iconfinder_103177_see_watch_view_eye_icon_512px;
            this.pictureBox1.Location = new System.Drawing.Point(181, 101);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(31, 22);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBox1_MouseDown);
            this.pictureBox1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pictureBox1_MouseUp);
            // 
            // chkFixDates
            // 
            this.chkFixDates.AutoSize = true;
            this.chkFixDates.Location = new System.Drawing.Point(14, 21);
            this.chkFixDates.Name = "chkFixDates";
            this.chkFixDates.Size = new System.Drawing.Size(253, 17);
            this.chkFixDates.TabIndex = 9;
            this.chkFixDates.Text = "Try Fix Dates for columns with headers including";
            this.chkFixDates.UseVisualStyleBackColor = true;
            this.chkFixDates.CheckedChanged += new System.EventHandler(this.chkFixDates_CheckedChanged);
            // 
            // chkIgnoreCols
            // 
            this.chkIgnoreCols.AutoSize = true;
            this.chkIgnoreCols.Location = new System.Drawing.Point(14, 64);
            this.chkIgnoreCols.Name = "chkIgnoreCols";
            this.chkIgnoreCols.Size = new System.Drawing.Size(237, 17);
            this.chkIgnoreCols.TabIndex = 9;
            this.chkIgnoreCols.Text = "Ignore Columns (e.g. Description, Comments)";
            this.chkIgnoreCols.UseVisualStyleBackColor = true;
            this.chkIgnoreCols.CheckedChanged += new System.EventHandler(this.chkIgnoreCols_CheckedChanged);
            // 
            // txtDateHeader
            // 
            this.txtDateHeader.Enabled = false;
            this.txtDateHeader.Location = new System.Drawing.Point(14, 38);
            this.txtDateHeader.Name = "txtDateHeader";
            this.txtDateHeader.Size = new System.Drawing.Size(251, 20);
            this.txtDateHeader.TabIndex = 6;
            // 
            // txtIgnoredCols
            // 
            this.txtIgnoredCols.Enabled = false;
            this.txtIgnoredCols.Location = new System.Drawing.Point(14, 81);
            this.txtIgnoredCols.Name = "txtIgnoredCols";
            this.txtIgnoredCols.Size = new System.Drawing.Size(251, 20);
            this.txtIgnoredCols.TabIndex = 10;
            this.txtIgnoredCols.Text = "Description,Comments";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkCleanHeaderNames);
            this.groupBox1.Controls.Add(this.chkFixDates);
            this.groupBox1.Controls.Add(this.txtIgnoredCols);
            this.groupBox1.Controls.Add(this.txtDateHeader);
            this.groupBox1.Controls.Add(this.chkIgnoreCols);
            this.groupBox1.Location = new System.Drawing.Point(224, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(273, 138);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // chkCleanHeaderNames
            // 
            this.chkCleanHeaderNames.AutoSize = true;
            this.chkCleanHeaderNames.Checked = true;
            this.chkCleanHeaderNames.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCleanHeaderNames.Location = new System.Drawing.Point(14, 115);
            this.chkCleanHeaderNames.Name = "chkCleanHeaderNames";
            this.chkCleanHeaderNames.Size = new System.Drawing.Size(127, 17);
            this.chkCleanHeaderNames.TabIndex = 11;
            this.chkCleanHeaderNames.Text = "Clean Header Names";
            this.chkCleanHeaderNames.UseVisualStyleBackColor = true;
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.linkLabel1.LinkColor = System.Drawing.Color.White;
            this.linkLabel1.Location = new System.Drawing.Point(462, 317);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(35, 13);
            this.linkLabel1.TabIndex = 12;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "About";
            this.linkLabel1.VisitedLinkColor = System.Drawing.Color.White;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lblColumns
            // 
            this.lblColumns.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblColumns.AutoSize = true;
            this.lblColumns.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblColumns.Location = new System.Drawing.Point(221, 157);
            this.lblColumns.Name = "lblColumns";
            this.lblColumns.Size = new System.Drawing.Size(54, 13);
            this.lblColumns.TabIndex = 14;
            this.lblColumns.Text = "Columns";
            // 
            // numHeaderOffset
            // 
            this.numHeaderOffset.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.numHeaderOffset.Location = new System.Drawing.Point(448, 155);
            this.numHeaderOffset.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numHeaderOffset.Name = "numHeaderOffset";
            this.numHeaderOffset.Size = new System.Drawing.Size(49, 20);
            this.numHeaderOffset.TabIndex = 15;
            this.numHeaderOffset.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numHeaderOffset.ValueChanged += new System.EventHandler(this.numHeaderOffset_ValueChanged);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(356, 157);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(76, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Header Offset:";
            // 
            // lvCols
            // 
            this.lvCols.ContextMenuStrip = this.contextMenueIgnoreDate;
            this.lvCols.FullRowSelect = true;
            this.lvCols.GridLines = true;
            this.lvCols.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.lvCols.HideSelection = false;
            this.lvCols.Location = new System.Drawing.Point(224, 178);
            this.lvCols.MultiSelect = false;
            this.lvCols.Name = "lvCols";
            this.lvCols.ShowGroups = false;
            this.lvCols.Size = new System.Drawing.Size(273, 95);
            this.lvCols.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.lvCols.TabIndex = 16;
            this.toolTip1.SetToolTip(this.lvCols, "Double Click to Ignore or mark as a Date");
            this.lvCols.UseCompatibleStateImageBehavior = false;
            this.lvCols.View = System.Windows.Forms.View.List;
            this.lvCols.SelectedIndexChanged += new System.EventHandler(this.lvCols_SelectedIndexChanged);
            this.lvCols.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.lvCols_MouseDoubleClick);
            // 
            // contextMenueIgnoreDate
            // 
            this.contextMenueIgnoreDate.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ignoreToolStripMenuItem,
            this.dateToolStripMenuItem});
            this.contextMenueIgnoreDate.Name = "contextMenueIgnoreDate";
            this.contextMenueIgnoreDate.Size = new System.Drawing.Size(109, 48);
            // 
            // ignoreToolStripMenuItem
            // 
            this.ignoreToolStripMenuItem.Name = "ignoreToolStripMenuItem";
            this.ignoreToolStripMenuItem.Size = new System.Drawing.Size(108, 22);
            this.ignoreToolStripMenuItem.Text = "&Ignore";
            this.ignoreToolStripMenuItem.Click += new System.EventHandler(this.ignoreToolStripMenuItem_Click);
            // 
            // dateToolStripMenuItem
            // 
            this.dateToolStripMenuItem.Name = "dateToolStripMenuItem";
            this.dateToolStripMenuItem.Size = new System.Drawing.Size(108, 22);
            this.dateToolStripMenuItem.Text = "&Date";
            this.dateToolStripMenuItem.Click += new System.EventHandler(this.dateToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.CadetBlue;
            this.ClientSize = new System.Drawing.Size(508, 348);
            this.Controls.Add(this.lvCols);
            this.Controls.Add(this.numHeaderOffset);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblColumns);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.passwordCheckBox);
            this.Controls.Add(this.pwdBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sheetsList);
            this.Controls.Add(this.chkOpenAfterFix);
            this.Controls.Add(this.fixButton);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Batch Fix Columns";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numHeaderOffset)).EndInit();
            this.contextMenueIgnoreDate.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.CheckBox chkOpenAfterFix;
        private System.Windows.Forms.Button fixButton;
        private System.Windows.Forms.ListBox sheetsList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox pwdBox;
        private System.Windows.Forms.CheckBox passwordCheckBox;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox chkFixDates;
        private System.Windows.Forms.CheckBox chkIgnoreCols;
        private System.Windows.Forms.TextBox txtDateHeader;
        private System.Windows.Forms.TextBox txtIgnoredCols;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label lblColumns;
        private System.Windows.Forms.NumericUpDown numHeaderOffset;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkCleanHeaderNames;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ContextMenuStrip contextMenueIgnoreDate;
        private System.Windows.Forms.ToolStripMenuItem ignoreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dateToolStripMenuItem;
        private System.Windows.Forms.ListView lvCols;
    }
}

