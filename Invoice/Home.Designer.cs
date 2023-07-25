﻿namespace Invoice
{
    partial class Home
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle11 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle4 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle5 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle6 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle7 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle8 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle9 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle10 = new DataGridViewCellStyle();
            BtnExport = new Button();
            label1 = new Label();
            TxtCompanyName = new TextBox();
            TxtEmail = new TextBox();
            label2 = new Label();
            TxtAddress = new TextBox();
            label3 = new Label();
            BtnSave = new Button();
            BtnDelete = new Button();
            DgvMain = new DataGridView();
            ColSrNo = new DataGridViewTextBoxColumn();
            ColDescription = new DataGridViewTextBoxColumn();
            ColHsn = new DataGridViewTextBoxColumn();
            ColGst = new DataGridViewTextBoxColumn();
            ColPcs = new DataGridViewTextBoxColumn();
            ColSalePrice = new DataGridViewTextBoxColumn();
            ColRate = new DataGridViewTextBoxColumn();
            ColDiscountP = new DataGridViewTextBoxColumn();
            ColAmount = new DataGridViewTextBoxColumn();
            panel1 = new Panel();
            label7 = new Label();
            TxtStateCode = new TextBox();
            CmbState = new ComboBox();
            label6 = new Label();
            label5 = new Label();
            DatePicker = new DateTimePicker();
            label4 = new Label();
            TxtInvoiceNo = new TextBox();
            panel2 = new Panel();
            MainPanel = new Panel();
            panel3 = new Panel();
            BtnReset = new Button();
            ((System.ComponentModel.ISupportInitialize)DgvMain).BeginInit();
            panel1.SuspendLayout();
            panel2.SuspendLayout();
            MainPanel.SuspendLayout();
            panel3.SuspendLayout();
            SuspendLayout();
            // 
            // BtnExport
            // 
            BtnExport.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnExport.Location = new Point(685, 8);
            BtnExport.Margin = new Padding(3, 2, 3, 2);
            BtnExport.Name = "BtnExport";
            BtnExport.Size = new Size(110, 34);
            BtnExport.TabIndex = 0;
            BtnExport.Text = "Export";
            BtnExport.UseVisualStyleBackColor = true;
            BtnExport.Click += BtnExport_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(350, 5);
            label1.Name = "label1";
            label1.Size = new Size(94, 15);
            label1.TabIndex = 1;
            label1.Text = "Company Name";
            // 
            // TxtCompanyName
            // 
            TxtCompanyName.Location = new Point(354, 22);
            TxtCompanyName.Margin = new Padding(3, 2, 3, 2);
            TxtCompanyName.Name = "TxtCompanyName";
            TxtCompanyName.Size = new Size(226, 23);
            TxtCompanyName.TabIndex = 1;
            // 
            // TxtEmail
            // 
            TxtEmail.Location = new Point(893, 22);
            TxtEmail.Margin = new Padding(3, 2, 3, 2);
            TxtEmail.Name = "TxtEmail";
            TxtEmail.Size = new Size(285, 23);
            TxtEmail.TabIndex = 3;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(889, 5);
            label2.Name = "label2";
            label2.Size = new Size(36, 15);
            label2.TabIndex = 3;
            label2.Text = "Email";
            // 
            // TxtAddress
            // 
            TxtAddress.Location = new Point(586, 22);
            TxtAddress.Margin = new Padding(3, 2, 3, 2);
            TxtAddress.Name = "TxtAddress";
            TxtAddress.Size = new Size(301, 23);
            TxtAddress.TabIndex = 2;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(586, 5);
            label3.Name = "label3";
            label3.Size = new Size(49, 15);
            label3.TabIndex = 5;
            label3.Text = "Address";
            // 
            // BtnSave
            // 
            BtnSave.BackColor = Color.Gainsboro;
            BtnSave.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnSave.Location = new Point(337, 8);
            BtnSave.Margin = new Padding(3, 2, 3, 2);
            BtnSave.Name = "BtnSave";
            BtnSave.Size = new Size(110, 34);
            BtnSave.TabIndex = 7;
            BtnSave.Text = "Save";
            BtnSave.UseVisualStyleBackColor = false;
            BtnSave.MouseLeave += BtnSave_MouseLeave;
            BtnSave.MouseHover += BtnSave_MouseHover;
            // 
            // BtnDelete
            // 
            BtnDelete.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnDelete.Location = new Point(453, 8);
            BtnDelete.Margin = new Padding(3, 2, 3, 2);
            BtnDelete.Name = "BtnDelete";
            BtnDelete.Size = new Size(110, 34);
            BtnDelete.TabIndex = 8;
            BtnDelete.Text = "Delete";
            BtnDelete.UseVisualStyleBackColor = true;
            BtnDelete.MouseLeave += BtnDelete_MouseLeave;
            BtnDelete.MouseHover += BtnDelete_MouseHover;
            // 
            // DgvMain
            // 
            DgvMain.AllowUserToAddRows = false;
            DgvMain.AllowUserToDeleteRows = false;
            DgvMain.AllowUserToResizeColumns = false;
            DgvMain.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = SystemColors.Control;
            dataGridViewCellStyle1.Font = new Font("Arial", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle1.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            DgvMain.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            DgvMain.ColumnHeadersHeight = 45;
            DgvMain.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            DgvMain.Columns.AddRange(new DataGridViewColumn[] { ColSrNo, ColDescription, ColHsn, ColGst, ColPcs, ColSalePrice, ColRate, ColDiscountP, ColAmount });
            dataGridViewCellStyle11.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle11.BackColor = SystemColors.Window;
            dataGridViewCellStyle11.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle11.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle11.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle11.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle11.WrapMode = DataGridViewTriState.False;
            DgvMain.DefaultCellStyle = dataGridViewCellStyle11;
            DgvMain.Location = new Point(2, 3);
            DgvMain.MultiSelect = false;
            DgvMain.Name = "DgvMain";
            DgvMain.RowHeadersVisible = false;
            DgvMain.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            DgvMain.RowTemplate.Height = 33;
            DgvMain.Size = new Size(1195, 428);
            DgvMain.TabIndex = 0;
            // 
            // ColSrNo
            // 
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle2.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            ColSrNo.DefaultCellStyle = dataGridViewCellStyle2;
            ColSrNo.HeaderText = "Sr No.";
            ColSrNo.Name = "ColSrNo";
            ColSrNo.ReadOnly = true;
            ColSrNo.Width = 80;
            // 
            // ColDescription
            // 
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
            ColDescription.DefaultCellStyle = dataGridViewCellStyle3;
            ColDescription.HeaderText = "Description Of Good";
            ColDescription.Name = "ColDescription";
            ColDescription.Width = 250;
            // 
            // ColHsn
            // 
            dataGridViewCellStyle4.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle4.Format = "N0";
            dataGridViewCellStyle4.NullValue = null;
            dataGridViewCellStyle4.WrapMode = DataGridViewTriState.False;
            ColHsn.DefaultCellStyle = dataGridViewCellStyle4;
            ColHsn.HeaderText = "HSN";
            ColHsn.Name = "ColHsn";
            // 
            // ColGst
            // 
            dataGridViewCellStyle5.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            ColGst.DefaultCellStyle = dataGridViewCellStyle5;
            ColGst.HeaderText = "GST";
            ColGst.Name = "ColGst";
            // 
            // ColPcs
            // 
            dataGridViewCellStyle6.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle6.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            ColPcs.DefaultCellStyle = dataGridViewCellStyle6;
            ColPcs.HeaderText = "Pcs";
            ColPcs.Name = "ColPcs";
            ColPcs.Width = 50;
            // 
            // ColSalePrice
            // 
            dataGridViewCellStyle7.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle7.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle7.Format = "N2";
            dataGridViewCellStyle7.NullValue = null;
            ColSalePrice.DefaultCellStyle = dataGridViewCellStyle7;
            ColSalePrice.HeaderText = "Sale Price";
            ColSalePrice.Name = "ColSalePrice";
            // 
            // ColRate
            // 
            dataGridViewCellStyle8.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle8.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle8.Format = "N2";
            ColRate.DefaultCellStyle = dataGridViewCellStyle8;
            ColRate.HeaderText = "Rate";
            ColRate.Name = "ColRate";
            // 
            // ColDiscountP
            // 
            dataGridViewCellStyle9.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle9.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle9.Format = "N2";
            ColDiscountP.DefaultCellStyle = dataGridViewCellStyle9;
            ColDiscountP.HeaderText = "Disc %";
            ColDiscountP.Name = "ColDiscountP";
            // 
            // ColAmount
            // 
            dataGridViewCellStyle10.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle10.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle10.Format = "N2";
            ColAmount.DefaultCellStyle = dataGridViewCellStyle10;
            ColAmount.HeaderText = "Amount";
            ColAmount.Name = "ColAmount";
            ColAmount.Width = 130;
            // 
            // panel1
            // 
            panel1.Controls.Add(label7);
            panel1.Controls.Add(TxtStateCode);
            panel1.Controls.Add(CmbState);
            panel1.Controls.Add(label6);
            panel1.Controls.Add(label5);
            panel1.Controls.Add(DatePicker);
            panel1.Controls.Add(label4);
            panel1.Controls.Add(TxtInvoiceNo);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(TxtCompanyName);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(TxtEmail);
            panel1.Controls.Add(TxtAddress);
            panel1.Controls.Add(label3);
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1200, 100);
            panel1.TabIndex = 10;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(161, 54);
            label7.Name = "label7";
            label7.Size = new Size(35, 15);
            label7.TabIndex = 13;
            label7.Text = "Code";
            // 
            // TxtStateCode
            // 
            TxtStateCode.Enabled = false;
            TxtStateCode.Location = new Point(158, 72);
            TxtStateCode.Margin = new Padding(3, 2, 3, 2);
            TxtStateCode.Name = "TxtStateCode";
            TxtStateCode.ReadOnly = true;
            TxtStateCode.Size = new Size(43, 23);
            TxtStateCode.TabIndex = 14;
            TxtStateCode.WordWrap = false;
            // 
            // CmbState
            // 
            CmbState.FormattingEnabled = true;
            CmbState.Location = new Point(4, 72);
            CmbState.Name = "CmbState";
            CmbState.Size = new Size(148, 23);
            CmbState.TabIndex = 4;
            CmbState.SelectedIndexChanged += CmbState_SelectedIndexChanged;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(7, 53);
            label6.Name = "label6";
            label6.Size = new Size(33, 15);
            label6.TabIndex = 11;
            label6.Text = "State";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(124, 5);
            label5.Name = "label5";
            label5.Size = new Size(31, 15);
            label5.TabIndex = 10;
            label5.Text = "Date";
            // 
            // DatePicker
            // 
            DatePicker.Location = new Point(124, 22);
            DatePicker.Name = "DatePicker";
            DatePicker.Size = new Size(220, 23);
            DatePicker.TabIndex = 9;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(8, 5);
            label4.Name = "label4";
            label4.Size = new Size(67, 15);
            label4.TabIndex = 7;
            label4.Text = "Invoice No.";
            // 
            // TxtInvoiceNo
            // 
            TxtInvoiceNo.Enabled = false;
            TxtInvoiceNo.Location = new Point(4, 22);
            TxtInvoiceNo.Margin = new Padding(3, 2, 3, 2);
            TxtInvoiceNo.Name = "TxtInvoiceNo";
            TxtInvoiceNo.ReadOnly = true;
            TxtInvoiceNo.Size = new Size(110, 23);
            TxtInvoiceNo.TabIndex = 8;
            // 
            // panel2
            // 
            panel2.Controls.Add(DgvMain);
            panel2.Location = new Point(0, 99);
            panel2.Name = "panel2";
            panel2.Size = new Size(1199, 433);
            panel2.TabIndex = 11;
            // 
            // MainPanel
            // 
            MainPanel.Controls.Add(panel3);
            MainPanel.Controls.Add(panel1);
            MainPanel.Controls.Add(panel2);
            MainPanel.Location = new Point(1, 2);
            MainPanel.Name = "MainPanel";
            MainPanel.Size = new Size(1201, 591);
            MainPanel.TabIndex = 12;
            // 
            // panel3
            // 
            panel3.Controls.Add(BtnReset);
            panel3.Controls.Add(BtnSave);
            panel3.Controls.Add(BtnDelete);
            panel3.Controls.Add(BtnExport);
            panel3.Location = new Point(3, 538);
            panel3.Name = "panel3";
            panel3.Size = new Size(1197, 51);
            panel3.TabIndex = 12;
            // 
            // BtnReset
            // 
            BtnReset.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnReset.Location = new Point(569, 8);
            BtnReset.Margin = new Padding(3, 2, 3, 2);
            BtnReset.Name = "BtnReset";
            BtnReset.Size = new Size(110, 34);
            BtnReset.TabIndex = 9;
            BtnReset.Text = "Reset";
            BtnReset.UseVisualStyleBackColor = true;
            BtnReset.Click += BtnReset_Click;
            // 
            // Home
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1204, 593);
            Controls.Add(MainPanel);
            Margin = new Padding(3, 2, 3, 2);
            Name = "Home";
            Text = "Home";
            Load += Home_Load;
            ((System.ComponentModel.ISupportInitialize)DgvMain).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            panel2.ResumeLayout(false);
            MainPanel.ResumeLayout(false);
            panel3.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button BtnExport;
        private Label label1;
        private TextBox TxtCompanyName;
        private TextBox TxtEmail;
        private Label label2;
        private TextBox TxtAddress;
        private Label label3;
        private Button BtnSave;
        private Button BtnDelete;
        private DataGridView DgvMain;
        private Panel panel1;
        private Panel panel2;
        private Panel MainPanel;
        private Panel panel3;
        private Label label4;
        private TextBox TxtInvoiceNo;
        private Label label5;
        private DateTimePicker DatePicker;
        private Label label7;
        private TextBox TxtStateCode;
        private ComboBox CmbState;
        private Label label6;
        private Button BtnReset;
        private DataGridViewTextBoxColumn ColSrNo;
        private DataGridViewTextBoxColumn ColDescription;
        private DataGridViewTextBoxColumn ColHsn;
        private DataGridViewTextBoxColumn ColGst;
        private DataGridViewTextBoxColumn ColPcs;
        private DataGridViewTextBoxColumn ColSalePrice;
        private DataGridViewTextBoxColumn ColRate;
        private DataGridViewTextBoxColumn ColDiscountP;
        private DataGridViewTextBoxColumn ColAmount;
    }
}