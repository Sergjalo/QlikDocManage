namespace WindowsFormsApplication4
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.dOF = new System.Windows.Forms.OpenFileDialog();
            this.lCnt = new System.Windows.Forms.Label();
            this.dgVars = new System.Windows.Forms.DataGridView();
            this.txValCorrect = new System.Windows.Forms.TextBox();
            this.txCommCorrect = new System.Windows.Forms.TextBox();
            this.txFlName = new System.Windows.Forms.TextBox();
            this.bSaveChg = new System.Windows.Forms.Button();
            this.col1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.bExport = new System.Windows.Forms.Button();
            this.bLoad = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txFlNameComm = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgVars)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(12, 15);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(93, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Get QV doc";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dOF
            // 
            this.dOF.FileName = "FleetCompCode.qvw";
            this.dOF.Filter = "QlikView|*.qvw|Excel|*.xls*";
            this.dOF.FilterIndex = 2;
            this.dOF.InitialDirectory = "d:\\work\\PNSK-CER\\";
            // 
            // lCnt
            // 
            this.lCnt.AutoSize = true;
            this.lCnt.Location = new System.Drawing.Point(12, 64);
            this.lCnt.Name = "lCnt";
            this.lCnt.Size = new System.Drawing.Size(0, 13);
            this.lCnt.TabIndex = 1;
            // 
            // dgVars
            // 
            this.dgVars.AllowUserToOrderColumns = true;
            this.dgVars.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgVars.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgVars.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.col1,
            this.col2,
            this.col3,
            this.col4});
            this.dgVars.Location = new System.Drawing.Point(15, 202);
            this.dgVars.Name = "dgVars";
            this.dgVars.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgVars.ShowCellToolTips = false;
            this.dgVars.Size = new System.Drawing.Size(1047, 349);
            this.dgVars.TabIndex = 3;
            this.dgVars.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgVars_CellContentClick);
            this.dgVars.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgVars_RowEnter);
            this.dgVars.RowLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgVars_RowLeave);
            this.dgVars.DoubleClick += new System.EventHandler(this.dgVars_DoubleClick);
            // 
            // txValCorrect
            // 
            this.txValCorrect.Location = new System.Drawing.Point(15, 96);
            this.txValCorrect.Multiline = true;
            this.txValCorrect.Name = "txValCorrect";
            this.txValCorrect.ReadOnly = true;
            this.txValCorrect.Size = new System.Drawing.Size(413, 100);
            this.txValCorrect.TabIndex = 4;
            // 
            // txCommCorrect
            // 
            this.txCommCorrect.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txCommCorrect.Location = new System.Drawing.Point(434, 96);
            this.txCommCorrect.Multiline = true;
            this.txCommCorrect.Name = "txCommCorrect";
            this.txCommCorrect.Size = new System.Drawing.Size(628, 100);
            this.txCommCorrect.TabIndex = 4;
            this.txCommCorrect.Leave += new System.EventHandler(this.txCommCorrect_Leave);
            // 
            // txFlName
            // 
            this.txFlName.Location = new System.Drawing.Point(111, 15);
            this.txFlName.Name = "txFlName";
            this.txFlName.Size = new System.Drawing.Size(508, 20);
            this.txFlName.TabIndex = 5;
            this.txFlName.Text = "d:\\work\\PNSK-CER\\FleetCompCode.qvw";
            // 
            // bSaveChg
            // 
            this.bSaveChg.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bSaveChg.Location = new System.Drawing.Point(12, 44);
            this.bSaveChg.Name = "bSaveChg";
            this.bSaveChg.Size = new System.Drawing.Size(93, 23);
            this.bSaveChg.TabIndex = 6;
            this.bSaveChg.Text = "Save changes";
            this.bSaveChg.UseVisualStyleBackColor = true;
            this.bSaveChg.Click += new System.EventHandler(this.bSaveChg_Click);
            // 
            // col1
            // 
            this.col1.HeaderText = "VarName";
            this.col1.Name = "col1";
            this.col1.ReadOnly = true;
            // 
            // col2
            // 
            this.col2.DividerWidth = 1;
            this.col2.HeaderText = "Value";
            this.col2.Name = "col2";
            this.col2.ReadOnly = true;
            this.col2.Width = 600;
            // 
            // col3
            // 
            this.col3.HeaderText = "Comment";
            this.col3.Name = "col3";
            this.col3.ReadOnly = true;
            this.col3.Width = 500;
            // 
            // col4
            // 
            this.col4.HeaderText = "VarIndex";
            this.col4.Name = "col4";
            this.col4.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 80);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Edit value:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(444, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(74, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Edit comment:";
            // 
            // bExport
            // 
            this.bExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bExport.Location = new System.Drawing.Point(19, 21);
            this.bExport.Name = "bExport";
            this.bExport.Size = new System.Drawing.Size(93, 23);
            this.bExport.TabIndex = 6;
            this.bExport.Text = "Export to xls";
            this.bExport.UseVisualStyleBackColor = true;
            this.bExport.Click += new System.EventHandler(this.bExport_Click);
            // 
            // bLoad
            // 
            this.bLoad.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bLoad.Location = new System.Drawing.Point(140, 21);
            this.bLoad.Name = "bLoad";
            this.bLoad.Size = new System.Drawing.Size(93, 23);
            this.bLoad.TabIndex = 6;
            this.bLoad.Text = "Load from xls";
            this.bLoad.UseVisualStyleBackColor = true;
            this.bLoad.Click += new System.EventHandler(this.bLoad_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.bLoad);
            this.groupBox1.Controls.Add(this.bExport);
            this.groupBox1.Controls.Add(this.txFlNameComm);
            this.groupBox1.Location = new System.Drawing.Point(625, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(437, 79);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Comments";
            // 
            // txFlNameComm
            // 
            this.txFlNameComm.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txFlNameComm.Location = new System.Drawing.Point(19, 53);
            this.txFlNameComm.Name = "txFlNameComm";
            this.txFlNameComm.Size = new System.Drawing.Size(411, 20);
            this.txFlNameComm.TabIndex = 5;
            this.txFlNameComm.Text = "d:\\work\\PNSK-CER\\VariablesMain.xls";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 563);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.bSaveChg);
            this.Controls.Add(this.txFlName);
            this.Controls.Add(this.txCommCorrect);
            this.Controls.Add(this.txValCorrect);
            this.Controls.Add(this.dgVars);
            this.Controls.Add(this.lCnt);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Variables from Qvw";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgVars)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog dOF;
        private System.Windows.Forms.Label lCnt;
        private System.Windows.Forms.DataGridView dgVars;
        private System.Windows.Forms.TextBox txValCorrect;
        private System.Windows.Forms.TextBox txCommCorrect;
        private System.Windows.Forms.TextBox txFlName;
        private System.Windows.Forms.Button bSaveChg;
        private System.Windows.Forms.DataGridViewTextBoxColumn col1;
        private System.Windows.Forms.DataGridViewTextBoxColumn col2;
        private System.Windows.Forms.DataGridViewTextBoxColumn col3;
        private System.Windows.Forms.DataGridViewTextBoxColumn col4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button bExport;
        private System.Windows.Forms.Button bLoad;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txFlNameComm;
    }
}

