namespace Invoice
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
            ChkIGST = new CheckBox();
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
            BtnExport.Location = new Point(783, 11);
            BtnExport.Name = "BtnExport";
            BtnExport.Size = new Size(126, 45);
            BtnExport.TabIndex = 0;
            BtnExport.Text = "&Export";
            BtnExport.UseVisualStyleBackColor = true;
            BtnExport.Click += BtnExport_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(400, 7);
            label1.Name = "label1";
            label1.Size = new Size(116, 20);
            label1.TabIndex = 1;
            label1.Text = "Company Name";
            // 
            // TxtCompanyName
            // 
            TxtCompanyName.Location = new Point(405, 29);
            TxtCompanyName.Name = "TxtCompanyName";
            TxtCompanyName.Size = new Size(258, 27);
            TxtCompanyName.TabIndex = 1;
            // 
            // TxtEmail
            // 
            TxtEmail.Location = new Point(1021, 29);
            TxtEmail.Name = "TxtEmail";
            TxtEmail.Size = new Size(325, 27);
            TxtEmail.TabIndex = 3;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(1016, 7);
            label2.Name = "label2";
            label2.Size = new Size(46, 20);
            label2.TabIndex = 3;
            label2.Text = "Email";
            // 
            // TxtAddress
            // 
            TxtAddress.Location = new Point(670, 29);
            TxtAddress.Name = "TxtAddress";
            TxtAddress.Size = new Size(343, 27);
            TxtAddress.TabIndex = 2;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(670, 7);
            label3.Name = "label3";
            label3.Size = new Size(62, 20);
            label3.TabIndex = 5;
            label3.Text = "Address";
            // 
            // BtnSave
            // 
            BtnSave.BackColor = Color.Gainsboro;
            BtnSave.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnSave.Location = new Point(385, 11);
            BtnSave.Name = "BtnSave";
            BtnSave.Size = new Size(126, 45);
            BtnSave.TabIndex = 7;
            BtnSave.Text = "Save";
            BtnSave.UseVisualStyleBackColor = false;
            BtnSave.Click += BtnSave_Click;
            BtnSave.MouseLeave += BtnSave_MouseLeave;
            BtnSave.MouseHover += BtnSave_MouseHover;
            // 
            // BtnDelete
            // 
            BtnDelete.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnDelete.Location = new Point(518, 11);
            BtnDelete.Name = "BtnDelete";
            BtnDelete.Size = new Size(126, 45);
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
            DgvMain.Location = new Point(2, 4);
            DgvMain.Margin = new Padding(3, 4, 3, 4);
            DgvMain.MultiSelect = false;
            DgvMain.Name = "DgvMain";
            DgvMain.RowHeadersVisible = false;
            DgvMain.RowHeadersWidth = 51;
            DgvMain.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            DgvMain.RowTemplate.Height = 33;
            DgvMain.Size = new Size(1366, 571);
            DgvMain.TabIndex = 0;
            DgvMain.CellValueChanged += DgvMain_CellValueChanged;
            DgvMain.KeyDown += DgvMain_KeyDown;
            // 
            // ColSrNo
            // 
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle2.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            ColSrNo.DefaultCellStyle = dataGridViewCellStyle2;
            ColSrNo.HeaderText = "Sr No.";
            ColSrNo.MinimumWidth = 6;
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
            ColDescription.MinimumWidth = 6;
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
            ColHsn.MinimumWidth = 6;
            ColHsn.Name = "ColHsn";
            ColHsn.Width = 125;
            // 
            // ColGst
            // 
            dataGridViewCellStyle5.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            ColGst.DefaultCellStyle = dataGridViewCellStyle5;
            ColGst.HeaderText = "GST";
            ColGst.MinimumWidth = 6;
            ColGst.Name = "ColGst";
            ColGst.Width = 125;
            // 
            // ColPcs
            // 
            dataGridViewCellStyle6.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle6.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            ColPcs.DefaultCellStyle = dataGridViewCellStyle6;
            ColPcs.HeaderText = "Pcs";
            ColPcs.MinimumWidth = 6;
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
            ColSalePrice.MinimumWidth = 6;
            ColSalePrice.Name = "ColSalePrice";
            ColSalePrice.Width = 125;
            // 
            // ColRate
            // 
            dataGridViewCellStyle8.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle8.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle8.Format = "N2";
            ColRate.DefaultCellStyle = dataGridViewCellStyle8;
            ColRate.HeaderText = "Rate";
            ColRate.MinimumWidth = 6;
            ColRate.Name = "ColRate";
            ColRate.Width = 125;
            // 
            // ColDiscountP
            // 
            dataGridViewCellStyle9.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle9.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle9.Format = "N2";
            ColDiscountP.DefaultCellStyle = dataGridViewCellStyle9;
            ColDiscountP.HeaderText = "Disc %";
            ColDiscountP.MinimumWidth = 6;
            ColDiscountP.Name = "ColDiscountP";
            ColDiscountP.Width = 125;
            // 
            // ColAmount
            // 
            dataGridViewCellStyle10.Alignment = DataGridViewContentAlignment.TopRight;
            dataGridViewCellStyle10.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle10.Format = "N2";
            ColAmount.DefaultCellStyle = dataGridViewCellStyle10;
            ColAmount.HeaderText = "Amount";
            ColAmount.MinimumWidth = 6;
            ColAmount.Name = "ColAmount";
            ColAmount.Width = 130;
            // 
            // panel1
            // 
            panel1.Controls.Add(ChkIGST);
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
            panel1.Margin = new Padding(3, 4, 3, 4);
            panel1.Name = "panel1";
            panel1.Size = new Size(1371, 133);
            panel1.TabIndex = 10;
            // 
            // ChkIGST
            // 
            ChkIGST.AutoSize = true;
            ChkIGST.Location = new Point(282, 97);
            ChkIGST.Name = "ChkIGST";
            ChkIGST.Size = new Size(61, 24);
            ChkIGST.TabIndex = 15;
            ChkIGST.Text = "IGST";
            ChkIGST.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(184, 72);
            label7.Name = "label7";
            label7.Size = new Size(44, 20);
            label7.TabIndex = 13;
            label7.Text = "Code";
            // 
            // TxtStateCode
            // 
            TxtStateCode.Enabled = false;
            TxtStateCode.Location = new Point(181, 96);
            TxtStateCode.Name = "TxtStateCode";
            TxtStateCode.ReadOnly = true;
            TxtStateCode.Size = new Size(49, 27);
            TxtStateCode.TabIndex = 14;
            TxtStateCode.WordWrap = false;
            // 
            // CmbState
            // 
            CmbState.FormattingEnabled = true;
            CmbState.Location = new Point(5, 96);
            CmbState.Margin = new Padding(3, 4, 3, 4);
            CmbState.Name = "CmbState";
            CmbState.Size = new Size(169, 28);
            CmbState.TabIndex = 4;
            CmbState.SelectedIndexChanged += CmbState_SelectedIndexChanged;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(8, 71);
            label6.Name = "label6";
            label6.Size = new Size(43, 20);
            label6.TabIndex = 11;
            label6.Text = "State";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(142, 7);
            label5.Name = "label5";
            label5.Size = new Size(41, 20);
            label5.TabIndex = 10;
            label5.Text = "Date";
            // 
            // DatePicker
            // 
            DatePicker.Location = new Point(142, 29);
            DatePicker.Margin = new Padding(3, 4, 3, 4);
            DatePicker.Name = "DatePicker";
            DatePicker.Size = new Size(251, 27);
            DatePicker.TabIndex = 9;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(9, 7);
            label4.Name = "label4";
            label4.Size = new Size(83, 20);
            label4.TabIndex = 7;
            label4.Text = "Invoice No.";
            // 
            // TxtInvoiceNo
            // 
            TxtInvoiceNo.Enabled = false;
            TxtInvoiceNo.Font = new Font("Arial", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
            TxtInvoiceNo.Location = new Point(5, 29);
            TxtInvoiceNo.Name = "TxtInvoiceNo";
            TxtInvoiceNo.ReadOnly = true;
            TxtInvoiceNo.Size = new Size(125, 26);
            TxtInvoiceNo.TabIndex = 8;
            TxtInvoiceNo.TextAlign = HorizontalAlignment.Right;
            // 
            // panel2
            // 
            panel2.Controls.Add(DgvMain);
            panel2.Location = new Point(0, 132);
            panel2.Margin = new Padding(3, 4, 3, 4);
            panel2.Name = "panel2";
            panel2.Size = new Size(1370, 577);
            panel2.TabIndex = 11;
            // 
            // MainPanel
            // 
            MainPanel.Controls.Add(panel3);
            MainPanel.Controls.Add(panel1);
            MainPanel.Controls.Add(panel2);
            MainPanel.Location = new Point(1, 3);
            MainPanel.Margin = new Padding(3, 4, 3, 4);
            MainPanel.Name = "MainPanel";
            MainPanel.Size = new Size(1373, 788);
            MainPanel.TabIndex = 12;
            // 
            // panel3
            // 
            panel3.Controls.Add(BtnReset);
            panel3.Controls.Add(BtnSave);
            panel3.Controls.Add(BtnDelete);
            panel3.Controls.Add(BtnExport);
            panel3.Location = new Point(3, 717);
            panel3.Margin = new Padding(3, 4, 3, 4);
            panel3.Name = "panel3";
            panel3.Size = new Size(1368, 68);
            panel3.TabIndex = 12;
            // 
            // BtnReset
            // 
            BtnReset.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            BtnReset.Location = new Point(650, 11);
            BtnReset.Name = "BtnReset";
            BtnReset.Size = new Size(126, 45);
            BtnReset.TabIndex = 9;
            BtnReset.Text = "Reset";
            BtnReset.UseVisualStyleBackColor = true;
            BtnReset.Click += BtnReset_Click;
            // 
            // Home
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1376, 791);
            Controls.Add(MainPanel);
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
        private CheckBox ChkIGST;
    }
}