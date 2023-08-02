using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
namespace Invoice
{

    public partial class Home : Form
    {

        public Home()
        {
            InitializeComponent();
        }

        //Class Variables
        Dictionary<string, string> states = new Dictionary<string, string>();


        private void Home_Load(object sender, EventArgs e)
        {
            //Adding Value in Dictionary
            StateDataLoad();


            TxtInvoiceNo.Text = DataReadWrite().ToString(); //Create A file in C Drive 


            //add Value and Properties To combobox 
            CmbState.Sorted = true;
            CmbState.AutoCompleteSource = AutoCompleteSource.ListItems;
            CmbState.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            CmbState.Items.AddRange(states.Keys.ToArray());

            //Data Grid View
            //Total Row
            DgvMain.Rows.Add();
            DgvMain.Rows[0].ReadOnly = true;
            //First Row
            DgvMain.Rows.Add();
            DgvMain[0, 1].Value = 1;
            DgvMain.Rows[1].Cells[1].Selected = true;
        }

        private void CmbState_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (CmbState.SelectedIndex != -1 && CmbState.SelectedItem != null)
                {
                    TxtStateCode.Text = states[CmbState.SelectedItem.ToString()];
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void BtnExport_Click(object sender, System.EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;
            try
            {
                int xRow = 1;
                int xCol = 1;
                int dataFillStartRow;
                int Srno, Description, Hsn, Gst, Pcs, SPrice, Rate, Per, Discount, Amount;
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                #region "My Work"
                oSheet.Cells[1, 10] = "TAX INVOICE";
                oSheet.get_Range("A1", "J1").Font.Bold = true;
                oSheet.get_Range("A1", "J1").Font.Size = 12;
                oSheet.get_Range("A1", "J1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A1", "J1").HorizontalAlignment =
                Excel.XlHAlign.xlHAlignCenter;
                oSheet.get_Range("A1", "J1").Merge();
                oSheet.Range["A1", "J1"].RowHeight = 24;

                //Seller Address
                xRow = 2;
                oSheet.Cells[xRow, 1] = "ANKIT STORE";
                oSheet.get_Range("A2", "c2").Font.Bold = true;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                //Invoice No

                oSheet.Cells[xRow, 4] = "Invoice No: ";
                oSheet.Cells[xRow + 1, 4] = TxtInvoiceNo.Text;
                oSheet.get_Range("D3", "F3").Font.Bold = true;
                oSheet.get_Range("D3", "F3").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                oSheet.get_Range("D2", "F3").Cells.BorderAround2();
                //Date
                DateTime dt = new DateTime();
                dt = DatePicker.Value;
                oSheet.Cells[xRow, 7] = "Dated ";
                oSheet.Cells[xRow + 1, 7] = dt.ToShortDateString();
                oSheet.get_Range("G3", "J3").Font.Bold = true;
                oSheet.get_Range("G3", "J3").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                oSheet.get_Range("G2", "J3").Cells.BorderAround2();
                oSheet.get_Range("G3", "J3").Merge();

                xRow = 3;
                oSheet.Cells[xRow, 1] = "71, Triveni Nagar 1st, Jaipur";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 4;
                oSheet.Cells[xRow, 1] = "Ph. - 8000262869";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                oSheet.Cells[xRow, 4] = "Delivery Note";
                oSheet.Cells[xRow, 7] = "Mode/Terms of Payment";
                oSheet.get_Range("D" + xRow.ToString(), "F" + (xRow + 1).ToString()).Cells.BorderAround2();

                oSheet.get_Range("D" + xRow.ToString(), "F" + xRow.ToString()).Merge();
                oSheet.get_Range("G" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                xRow = 5;
                oSheet.Cells[xRow, 1] = "GSTIN/UIN: 08HLVPD9817B1ZO";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 6;
                oSheet.Cells[xRow, 1] = "State Name: Rajasthan, Code : 08";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                oSheet.Cells[xRow, 4] = "Supplier's Ref.";
                oSheet.Cells[xRow, 7] = "Other Reference(s)";
                oSheet.get_Range("G" + xRow.ToString(), "J" + (xRow + 1).ToString()).Cells.BorderAround2();
                oSheet.get_Range("D" + xRow.ToString(), "F" + xRow.ToString()).Merge();
                oSheet.get_Range("G" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                xRow = 7;
                oSheet.Cells[xRow, 1] = "E - Mail: ankitramola101 @gmail.com";


                //lable
                xRow = 8;
                oSheet.Cells[xRow, 1] = "Bayer";
                oSheet.Cells[xRow, 4] = "Buyer's Order No.";
                oSheet.Cells[xRow, 7] = "Dated";
                oSheet.get_Range("D" + xRow.ToString(), "F" + (xRow + 1).ToString()).Cells.BorderAround2();
                oSheet.get_Range("D" + xRow.ToString(), "F" + xRow.ToString()).Merge();
                oSheet.get_Range("G" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                //Buyer Address
                xRow = 9;
                oSheet.Cells[xRow, 1] = TxtCompanyName.Text;
                oSheet.get_Range("A9", "c9").Font.Bold = true;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();


                xRow = 10;
                oSheet.Cells[xRow, 1] = TxtAddress.Text;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();
                oSheet.get_Range("A8", "C8").Cells.BorderAround2();

                oSheet.Cells[xRow, 4] = "Despatch Document No.";
                oSheet.Cells[xRow, 7] = "Delivery Note Date";
                oSheet.get_Range("G" + xRow.ToString(), "J" + (xRow + 1).ToString()).Cells.BorderAround2();
                oSheet.get_Range("D" + xRow.ToString(), "F" + xRow.ToString()).Merge();
                oSheet.get_Range("G" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                xRow = 11;
                oSheet.Cells[xRow, 1] = "E - Mail: " + TxtEmail.Text;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 12;
                oSheet.Cells[xRow, 1] = "State : " + CmbState.Text.ToString() + ", Code : " + TxtStateCode.Text;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();
                //oSheet.get_Range("A9", "C13").Cells.BorderAround2();

                oSheet.Cells[xRow, 4] = "Despatched through";
                oSheet.Cells[xRow, 7] = "Destination";
                oSheet.get_Range("D" + xRow.ToString(), "F" + (xRow + 1).ToString()).Cells.BorderAround2();
                oSheet.get_Range("G" + xRow.ToString(), "J" + (xRow + 3).ToString()).Cells.BorderAround2();
                oSheet.get_Range("D" + xRow.ToString(), "F" + xRow.ToString()).Merge();
                oSheet.get_Range("G" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                xRow = 13;
                oSheet.Cells[xRow + 1, 4] = "Terms of Delivery";
                oSheet.get_Range("D" + (xRow + 1).ToString(), "F" + (xRow + 2).ToString()).Cells.BorderAround2();
                oSheet.get_Range("D" + (xRow + 1).ToString(), "F" + (xRow + 1).ToString()).Merge();
                xRow = 15;
                oSheet.get_Range("A2", "C15").Cells.BorderAround2();
                oSheet.get_Range("A2", "J15").Cells.BorderAround2();

                //Grid Column
                xRow += 1;

                Srno = xCol;
                oSheet.Cells[xRow, Srno] = "Sr No";

                xCol += 1;
                Description = xCol;
                oSheet.Cells[xRow, Description] = "Description of Goods";

                xCol += 1;
                Hsn = xCol;
                oSheet.Cells[xRow, Hsn] = "HSN/SAC";

                xCol += 1;
                Gst = xCol;
                oSheet.Cells[xRow, Gst] = "GST";

                xCol += 1;
                Pcs = xCol;
                oSheet.Cells[xRow, Pcs] = "Pcs";

                xCol += 1;
                SPrice = xCol;
                oSheet.Cells[xRow, SPrice] = "Sale Price";

                xCol += 1;
                Rate = xCol;
                oSheet.Cells[xRow, Rate] = "Rate";

                xCol += 1;
                Per = xCol;
                oSheet.Cells[xRow, Per] = "per";

                xCol += 1;
                Discount = xCol;
                oSheet.Cells[xRow, Discount] = "Disc %";

                xCol += 1;
                Amount = xCol;
                oSheet.Cells[xRow, Amount] = "Amount";
                oSheet.get_Range("A16", "J16").Cells.BorderAround2();

                //Data Filling
                xRow += 1;
                dataFillStartRow = xRow; //WIll require later
                float Igst = 0, Sgst = 0;
                foreach (DataGridViewRow dgRow in DgvMain.Rows)
                {
                    if (dgRow.Index > 0)
                    {
                        oSheet.Cells[xRow, Srno] = dgRow.Cells[ColSrNo.Index].Value;
                        oSheet.Cells[xRow, Description] = dgRow.Cells[ColDescription.Index].Value;
                        oSheet.Cells[xRow, Hsn] = dgRow.Cells[ColHsn.Index].Value;
                        oSheet.Cells[xRow, Gst] = isNull(dgRow.Cells[ColGst.Index].Value) ? 0 : dgRow.Cells[ColGst.Index].Value.ToString() + "%";
                        oSheet.Cells[xRow, Pcs] = isNull(dgRow.Cells[ColPcs.Index].Value) ? 0 : dgRow.Cells[ColPcs.Index].Value.ToString();
                        oSheet.Cells[xRow, SPrice] = dgRow.Cells[ColSalePrice.Index].Value;
                        oSheet.Cells[xRow, Rate] = dgRow.Cells[ColRate.Index].Value;
                        oSheet.Cells[xRow, Per] = "";
                        oSheet.Cells[xRow, Discount] = isNull(dgRow.Cells[ColDiscountP.Index].Value) ? 0 : dgRow.Cells[ColDiscountP.Index].Value.ToString() + "%";

                        oSheet.Cells[xRow, Amount] = dgRow.Cells[ColAmount.Index].Value;

                        if (!ChkIGST.Checked)
                        {
                            Sgst += float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 200;
                        }
                        else
                        {
                            Igst += float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 100;
                        }

                        xRow += 1;

                    }
                }

                //SGST OR IGST
                if (!ChkIGST.Checked)
                {
                    xRow += 1;
                    oSheet.Cells[xRow, Description] = "SGST";
                    oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    oSheet.Cells[xRow, Amount] = Sgst;
                    xRow += 1;
                    oSheet.Cells[xRow, Description] = "CGST";
                    oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    oSheet.Cells[xRow, Amount] = Sgst;
                }
                else
                {
                    xRow += 1;
                    oSheet.Cells[xRow, Description] = "IGST";
                    oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    oSheet.Cells[xRow, Amount] = Igst;
                }

                xRow += 1;
                oSheet.Cells[xRow, Description] = "Round Off(+)";
                oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;


                //Total Calculation
                xRow += 2;
                //Total Amount
                oSheet.Cells[xRow, Description] = "Total";
                oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                oSheet.Cells[xRow, Amount] = "=SUM(" + "J" + dataFillStartRow.ToString() + ":" + "J" + (xRow - 1).ToString() + ")";
                //Total PCs
                oSheet.Cells[xRow, Pcs] = "=SUM(" + "E" + dataFillStartRow.ToString() + ":" + "E" + (xRow - 1).ToString() + ")";

                oSheet.Range["F" + dataFillStartRow, "G" + xRow].NumberFormatLocal = "0.00";
                oSheet.Range["J" + dataFillStartRow, "J" + xRow].NumberFormatLocal = "0.00";

                oSheet.get_Range("A" + xRow.ToString(), "J" + xRow.ToString()).Font.Bold = true;
                oSheet.get_Range("J" + dataFillStartRow.ToString(), "J" + xRow.ToString()).Font.Bold = true;
                oSheet.get_Range("B" + dataFillStartRow.ToString(), "B" + xRow.ToString()).Font.Bold = true;
                oSheet.get_Range("A" + (xRow).ToString(), "J" + (xRow).ToString()).Cells.BorderAround2();


                oSheet.get_Range("A" + (dataFillStartRow - 1).ToString(), "A" + (xRow).ToString()).Cells.BorderAround2();
                oSheet.get_Range("C" + (dataFillStartRow - 1).ToString(), "C" + (xRow).ToString()).Cells.BorderAround2();
                oSheet.get_Range("E" + (dataFillStartRow - 1).ToString(), "E" + (xRow).ToString()).Cells.BorderAround2();
                oSheet.get_Range("G" + (dataFillStartRow - 1).ToString(), "G" + (xRow).ToString()).Cells.BorderAround2();
                oSheet.get_Range("I" + (dataFillStartRow - 1).ToString(), "I" + (xRow).ToString()).Cells.BorderAround2();
                oSheet.get_Range("J" + (dataFillStartRow - 1).ToString(), "j" + (xRow).ToString()).Cells.BorderAround2();
                oSheet.get_Range("A" + (dataFillStartRow - 1).ToString(), "J" + (xRow).ToString()).Cells.BorderAround2();

                //Formating Above data
                xRow += 1;
                oSheet.Cells[xRow, Srno] = "Amount Chargeable (in words)";

                xRow += 1;
                oSheet.Cells[xRow, Srno] = CommonFunctions.ConvertToWords(oSheet.Cells[xRow - 2, Amount].Text);

                oSheet.get_Range("A" + (xRow).ToString(), "J" + (xRow).ToString()).Font.Bold = true;
                oSheet.get_Range("A" + (xRow - 1).ToString(), "J" + (xRow).ToString()).Cells.BorderAround2();

                xRow += 1;
                oSheet.Cells[xRow, Srno] = "HSN/SAC";
                oSheet.get_Range("A" + xRow.ToString(), "C" + (xRow + 2).ToString()).Merge();

                oSheet.Cells[xRow, Gst] = "Taxable Value";
                oSheet.get_Range("D" + xRow.ToString(), "D" + (xRow + 2).ToString()).Merge();
                oSheet.get_Range("D" + xRow.ToString(), "D" + (xRow + 2).ToString()).WrapText = true;

                //Central tax
                if (!ChkIGST.Checked)
                {
                    oSheet.Cells[xRow, Pcs] = "Central Tax";
                    oSheet.get_Range("E" + xRow.ToString(), "F" + (xRow).ToString()).Merge();

                    oSheet.Cells[xRow + 1, Pcs] = "Rate";
                    oSheet.Cells[xRow + 1, SPrice] = "Amount";
                    oSheet.get_Range("E" + (xRow + 1).ToString(), "E" + (xRow + 2).ToString()).Merge();
                    oSheet.get_Range("F" + (xRow + 1).ToString(), "F" + (xRow + 2).ToString()).Merge();


                    oSheet.Cells[xRow, Rate] = "State Tax";
                    oSheet.get_Range("G" + xRow.ToString(), "I" + (xRow).ToString()).Merge();


                    oSheet.Cells[xRow + 1, Rate] = "Rate";
                    oSheet.Cells[xRow + 1, Per] = "Amount";
                    oSheet.get_Range("G" + (xRow + 1).ToString(), "G" + (xRow + 2).ToString()).Merge();
                    oSheet.get_Range("H" + (xRow + 1).ToString(), "I" + (xRow + 2).ToString()).Merge();

                }
                else
                {
                    oSheet.Cells[xRow, Pcs] = "Inter State Tax";
                    oSheet.get_Range("E" + xRow.ToString(), "I" + (xRow).ToString()).Merge();
                    oSheet.Cells[xRow + 1, Pcs] = "Rate";
                    oSheet.Cells[xRow + 1, Rate] = "Amount";
                    oSheet.get_Range("E" + (xRow + 1).ToString(), "F" + (xRow + 2).ToString()).Merge();
                    oSheet.get_Range("G" + (xRow + 1).ToString(), "I" + (xRow + 2).ToString()).Merge();
                    oSheet.get_Range("G" + (xRow).ToString(), "I" + (xRow).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;
                }

                //Total Tax
                oSheet.Cells[xRow, Amount] = "Total Tax Amount";
                oSheet.get_Range("J" + xRow.ToString(), "J" + (xRow + 2).ToString()).Merge();
                oSheet.get_Range("J" + xRow.ToString(), "J" + (xRow + 2).ToString()).WrapText = true;


                oSheet.get_Range("A" + xRow.ToString(), "J" + (xRow + 2).ToString()).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A" + xRow.ToString(), "J" + (xRow + 2).ToString()).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //Tax Calculation
                xRow += 3;
                int taxStartRow = xRow;
                foreach (DataGridViewRow dgRow in DgvMain.Rows)
                {
                    if (dgRow.Index > 0)
                    {
                        oSheet.Cells[xRow, Srno] = dgRow.Cells[ColHsn.Index].Value;
                        oSheet.get_Range("A" + xRow.ToString(), "C" + (xRow).ToString()).Merge();
                        oSheet.get_Range("A" + (xRow + 1).ToString(), "C" + (xRow + 1).ToString()).Merge();

                        oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        oSheet.Cells[xRow, Gst] = dgRow.Cells[ColAmount.Index].Value;
                        if (!ChkIGST.Checked)
                        {
                            //central
                            oSheet.Cells[xRow, Pcs] = float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 2;
                            oSheet.Cells[xRow, SPrice] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 200;
                            //state
                            oSheet.Cells[xRow, Rate] = float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 2;
                            oSheet.Cells[xRow, Per] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 200;
                            oSheet.get_Range("H" + xRow.ToString(), "I" + (xRow).ToString()).Merge();
                            oSheet.get_Range("H" + (xRow + 1).ToString(), "I" + (xRow + 1).ToString()).Merge();
                        }
                        else
                        {
                            oSheet.Cells[xRow, Pcs] = dgRow.Cells[ColGst.Index].Value;
                            oSheet.Cells[xRow, Per] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 100;
                            oSheet.get_Range("G" + xRow.ToString(), "I" + (xRow).ToString()).Merge();
                            oSheet.get_Range("G" + (xRow + 1).ToString(), "I" + (xRow + 1).ToString()).Merge();
                            oSheet.get_Range("E" + xRow.ToString(), "F" + (xRow).ToString()).Merge();
                            oSheet.get_Range("E" + (xRow + 1).ToString(), "F" + (xRow + 1).ToString()).Merge();
                        }

                        //total
                        oSheet.Cells[xRow, Amount] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 100;

                        xRow += 1;
                    }
                }

                //Tax Total 
                oSheet.Cells[xRow, Srno] = "Total";

                oSheet.Cells[xRow, Gst] = "=SUM(" + "D" + taxStartRow.ToString() + ":" + "D" + (xRow - 1).ToString() + ")";

                if (!ChkIGST.Checked)
                {
                    oSheet.Cells[xRow, SPrice] = "=SUM(" + "F" + taxStartRow.ToString() + ":" + "F" + (xRow - 1).ToString() + ")";
                    oSheet.Cells[xRow, Per] = "=SUM(" + "H" + taxStartRow.ToString() + ":" + "H" + (xRow - 1).ToString() + ")";

                }
                else
                {
                    oSheet.Cells[xRow, Rate] = "=SUM(" + "G" + taxStartRow.ToString() + ":" + "G" + (xRow - 1).ToString() + ")";
                }

                oSheet.Cells[xRow, Amount] = "=SUM(" + "J" + taxStartRow.ToString() + ":" + "J" + (xRow - 1).ToString() + ")";

                oSheet.Range["D" + taxStartRow, "J" + xRow].NumberFormatLocal = "0.00";
                oSheet.get_Range("A" + xRow.ToString(), "J" + xRow.ToString()).Font.Bold = true;
                oSheet.get_Range("A" + (taxStartRow - 3).ToString(), "J" + (xRow).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;

                //Disclamer
                xRow += 1;
                oSheet.get_Range("A" + (xRow).ToString(), "J" + (xRow + 9).ToString()).Cells.BorderAround2();


                oSheet.Cells[xRow, Srno] = "Tax Amount (in words) :" + CommonFunctions.ConvertToWords(oSheet.Cells[xRow - 1, Amount].Text);
                xRow += 1;
                oSheet.Cells[xRow, Srno] = "company's PAN : " + "HLVPD9817B";
                //oSheet.Cells[xRow, Description] = "HLVPD9817B";
                oSheet.get_Range("B" + xRow.ToString(), "B" + xRow.ToString()).Font.Bold = true;
                xRow++;
                oSheet.Cells[xRow, Srno] = "Declaration";
                xRow++;
                oSheet.Cells[xRow, Srno] = "Remarks :1. Goods Once sold will not be taken back.";
                xRow++;
                oSheet.Cells[xRow, Srno] = "2. Interest @24% will be charged if delayed beyond payment term.";
                xRow++;
                oSheet.Cells[xRow, Srno] = "3. Bounced cheque will attract Rs.500/- penalty";
                xRow++;
                oSheet.Cells[xRow, Srno] = "4. All warranty claims are Subject to the terms laid down by our principal.";
                xRow++;
                oSheet.Cells[xRow, Srno] = "5. WARRANTY VALID BY COMPANY SERVICE CENTER.";
                xRow++;
                oSheet.Cells[xRow, Srno] = "INDUSIND BANK IFSC CODE INDB0000707 A/C NO. 255525555555";

                #endregion

                //AutoFit columns A:D.
                // oRng = oSheet.get_Range("A1", "J30");
                //oRng.EntireColumn.AutoFit();

                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }


        private void BtnSave_MouseHover(object sender, EventArgs e)
        {

            BtnSave.BackColor = System.Drawing.Color.Green;
            BtnSave.ForeColor = System.Drawing.Color.White;
        }

        private void BtnSave_MouseLeave(object sender, EventArgs e)
        {
            BtnSave.BackColor = System.Drawing.Color.Gainsboro;
            BtnSave.ForeColor = System.Drawing.Color.Black;
        }

        private void BtnDelete_MouseHover(object sender, EventArgs e)
        {
            BtnDelete.BackColor = System.Drawing.Color.DarkRed;
            BtnDelete.ForeColor = System.Drawing.Color.White;
        }

        private void BtnDelete_MouseLeave(object sender, EventArgs e)
        {
            BtnDelete.BackColor = System.Drawing.Color.Gainsboro;
            BtnDelete.ForeColor = System.Drawing.Color.Black;
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            TxtAddress.Text = "";
            TxtCompanyName.Text = "";
            TxtEmail.Text = "";
            DatePicker.ResetText();

            DgvMain.Rows.Clear();
            DgvMain.Rows.Add();
            DgvMain.Rows[0].ReadOnly = true;
            //First Row
            DgvMain.Rows.Add();
            DgvMain[0, 1].Value = 1;
            DgvMain.Rows[1].Cells[1].Selected = true;
        }

        private void DgvMain_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Down)
            {
                if (DgvMain.CurrentCell.RowIndex == DgvMain.Rows.Count - 1)
                {
                    int rowCount = DgvMain.Rows.Count;

                    DgvMain.Rows.Add();
                    DgvMain[0, rowCount].Value = rowCount;
                }
            }
            else if (e.KeyCode == Keys.Delete)
            {
                int rowIndex = DgvMain.CurrentCell.RowIndex;

                if (rowIndex > 0)
                {

                    if (DialogResult.Yes == MessageBox.Show("Sure You Want To Delete This Row??", "Row Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                    {

                        if (DgvMain.Rows.Count - 1 == rowIndex && rowIndex == 1)
                        {

                            DgvMain.Rows.RemoveAt(rowIndex);
                            DgvMain.Rows.Add();
                            DgvMain[0, 1].Value = 1;
                            DgvMain.Rows[1].Cells[1].Selected = true;
                        }
                        else
                        {
                            DgvMain.Rows.RemoveAt(rowIndex);
                        }
                    }

                }
            }

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {

        }

        private void StateDataLoad()
        {
            states.Add("Andhra Pradesh", "37");
            states.Add("Arunachal Pradesh", "12");
            states.Add("Assam", "18");
            states.Add("Bihar", "10");
            states.Add("Chattisgarh", "22");
            states.Add("Delhi", "07");
            states.Add("Goa", "30");
            states.Add("Gujarat", "24");
            states.Add("Haryana", "06");
            states.Add("Himachal Pradesh", "02");
            states.Add("Jammu and Kashmir", "01");
            states.Add("Jharkhand", "20");
            states.Add("Karnataka", "29");
            states.Add("Kerala", "32");
            states.Add("Lakshadweep Islands", "31");
            states.Add("Madhya Pradesh", "23");
            states.Add("Maharashtra", "27");
            states.Add("Manipur", "14");
            states.Add("Meghalaya", "17");
            states.Add("Mizoram", "15");
            states.Add("Nagaland", "13");
            states.Add("Odisha", "21");
            states.Add("Pondicherry", "34");
            states.Add("Punjab", "03");
            states.Add("Rajasthan", "08");
            states.Add("Sikkim", "11");
            states.Add("Tamil Nadu", "33");
            states.Add("Telangana", "36");
            states.Add("Tripura", "16");
            states.Add("Uttar Pradesh", "09");
            states.Add("Uttarakhand", "05");
            states.Add("West Bengal", "19");
        }

        private Boolean isNull(Object x)
        {
            if (x == null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private Int32 DataReadWrite()
        {
            string directoryPath = @"C:\Invoice\";
            string fileName = "InvoiceNo.txt";
            string fullPath = Path.Combine(directoryPath, fileName);

            try
            {
                // Check if the directory exists, if not, create it
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                    // Now, you can access the file or perform any operations on it
                    // For example, writing some data to the file:
                    File.WriteAllText(fullPath, "InvoiceNo:105");
                    Console.WriteLine("Data written to the file successfully.");

                    return 105;

                }
                else
                {
                    string data = File.ReadAllText(fullPath);
                    return Int32.Parse(data.Substring(10));
                }


            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine("Error: The specified directory does not exist.");
                Console.WriteLine(ex.Message);
                return 1;
            }
            catch (Exception ex)
            {
                // Handle other exceptions here if needed
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                return 1;
            }
        }

        private void DgvMain_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > 0)
            {
                if (e.ColumnIndex == ColRate.Index)
                {
                    if (Int16.Parse((string)DgvMain.Rows[e.RowIndex].Cells[ColPcs.Index].Value) > 0)
                    {
                        DgvMain.Rows[e.RowIndex].Cells[ColAmount.Index].Value =( float.Parse((string)DgvMain.Rows[e.RowIndex].Cells[ColRate.Index].Value) * Int16.Parse((string)DgvMain.Rows[e.RowIndex].Cells[ColPcs.Index].Value)).ToString();
                    }
                }
            }
        }
    }
}