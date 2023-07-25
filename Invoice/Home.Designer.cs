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
            Export = new Button();
            label1 = new Label();
            TxtCompanyName = new TextBox();
            TxtEmail = new TextBox();
            label2 = new Label();
            TxtAddress = new TextBox();
            label3 = new Label();
            button1 = new Button();
            button2 = new Button();
            dataGridView1 = new DataGridView();
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
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            panel1.SuspendLayout();
            panel2.SuspendLayout();
            MainPanel.SuspendLayout();
            panel3.SuspendLayout();
            SuspendLayout();
            // 
            // Export
            // 
            Export.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            Export.Location = new Point(548, 8);
            Export.Margin = new Padding(3, 2, 3, 2);
            Export.Name = "Export";
            Export.Size = new Size(110, 34);
            Export.TabIndex = 0;
            Export.Text = "Export";
            Export.UseVisualStyleBackColor = true;
            Export.Click += button1_Click;
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
            // button1
            // 
            button1.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            button1.Location = new Point(316, 8);
            button1.Margin = new Padding(3, 2, 3, 2);
            button1.Name = "button1";
            button1.Size = new Size(110, 34);
            button1.TabIndex = 7;
            button1.Text = "button1";
            button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            button2.Font = new Font("Arial", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            button2.Location = new Point(432, 8);
            button2.Margin = new Padding(3, 2, 3, 2);
            button2.Name = "button2";
            button2.Size = new Size(110, 34);
            button2.TabIndex = 8;
            button2.Text = "button2";
            button2.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(2, 3);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(1195, 428);
            dataGridView1.TabIndex = 9;
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
            panel2.Controls.Add(dataGridView1);
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
            panel3.Controls.Add(button1);
            panel3.Controls.Add(button2);
            panel3.Controls.Add(Export);
            panel3.Location = new Point(3, 538);
            panel3.Name = "panel3";
            panel3.Size = new Size(1197, 51);
            panel3.TabIndex = 12;
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
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            panel2.ResumeLayout(false);
            MainPanel.ResumeLayout(false);
            panel3.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button Export;
        private Label label1;
        private TextBox TxtCompanyName;
        private TextBox TxtEmail;
        private Label label2;
        private TextBox TxtAddress;
        private Label label3;
        private Button button1;
        private Button button2;
        private DataGridView dataGridView1;
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
    }
}