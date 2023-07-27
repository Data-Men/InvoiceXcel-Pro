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
            try{
                int xRow = 1;
                int xCol = 1;
                int dataFillStartRow;
                int Srno, Description, Hsn, Gst, Pcs, SPrice, Rate, Per,Discount, Amount;
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
                oSheet.get_Range("A1", "J1").HorizontalAlignment=
                Excel.XlHAlign.xlHAlignCenter;
                oSheet.get_Range("A1", "J1").Merge();
                oSheet.Range["A1", "J1"].RowHeight = 24;

                //Seller Address
                xRow = 2;
                oSheet.Cells[xRow, 1] = "ANKIT STORE";
                oSheet.get_Range("A2", "c2").Font.Bold = true;
                oSheet.get_Range("A"+xRow.ToString(), "C"+xRow.ToString()).Merge();

                //Invoice No
                oSheet.Cells[xRow, 4] = "Invoice No: "+TxtInvoiceNo.Text;
                oSheet.get_Range("D"+xRow.ToString(), "F"+(xRow+1).ToString()).Merge();
                oSheet.get_Range("D2", "F3").HorizontalAlignment =
                Excel.XlHAlign.xlHAlignCenter;
                oSheet.get_Range("D2", "F3").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;
                //Date
                DateTime dt = new DateTime();
                dt = DatePicker.Value;
                oSheet.Cells[xRow, 7] = "Dated " + dt.ToShortDateString(); 
                oSheet.get_Range("G" + xRow.ToString(), "J" + (xRow+1).ToString()).Merge();
                oSheet.get_Range("G2", "J3").HorizontalAlignment =
               Excel.XlHAlign.xlHAlignCenter;
                oSheet.get_Range("G2", "J3").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;
                xRow = 3;
                oSheet.Cells[xRow, 1] = "71, Triveni Nagar 1st, Jaipur";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 4;
                oSheet.Cells[xRow, 1] = "Ph. - 8000262869";
                oSheet.get_Range("A"+xRow.ToString(), "C"+xRow.ToString()).Merge();

                xRow = 5;
                oSheet.Cells[xRow, 1] = "GSTIN/UIN: 08HLVPD9817B1ZO";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 6;
                oSheet.Cells[xRow, 1] = "State Name: Rajasthan, Code : 08";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 7;
                oSheet.Cells[xRow, 1] = "E - Mail: ankitramola101 @gmail.com";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                //oSheet.get_Range("A2", "C7").Merge();
                //lable
                xRow = 8;
                oSheet.Cells[xRow, 1] = "Bayer";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                //Buyer Address
                xRow = 9;
                oSheet.Cells[xRow, 1] = TxtCompanyName.Text;
                oSheet.get_Range("A9", "c9").Font.Bold = true;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 10;
                oSheet.Cells[xRow, 1] = TxtAddress.Text;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 11;
                oSheet.Cells[xRow, 1] = "E - Mail: "+ TxtEmail.Text ;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 12;
                oSheet.Cells[xRow, 1] = "State : "+CmbState.Text.ToString()+", Code : "+TxtStateCode.Text;
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                xRow = 13;
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
                Gst= xCol;
                oSheet.Cells[xRow, Gst] = "GST";

                xCol += 1;
                Pcs = xCol;
                oSheet.Cells[xRow, Pcs] = "Pcs";

                xCol += 1;
                SPrice= xCol;
                oSheet.Cells[xRow, SPrice] = "Sale Price";

                xCol += 1;
                Rate= xCol;
                oSheet.Cells[xRow, Rate] = "Rate";

                xCol += 1;
                Per = xCol;
                oSheet.Cells[xRow, Per] = "per";

                xCol += 1;
                Discount = xCol;
                oSheet.Cells[xRow, Discount] = "Disc %";

                xCol += 1;
                Amount= xCol;
                oSheet.Cells[xRow, Amount] = "Amount";

                //Data Filling
                xRow+= 1;
                dataFillStartRow = xRow; //WIll require later
                float Igst=0, Sgst=0;
                foreach (DataGridViewRow dgRow   in DgvMain.Rows)
                {
                    if (dgRow.Index>0)
                    {
                        oSheet.Cells[xRow, Srno] = dgRow.Cells[ColSrNo.Index].Value;
                        oSheet.Cells[xRow, Description] = dgRow.Cells[ColDescription.Index].Value;
                        oSheet.Cells[xRow, Hsn] = dgRow.Cells[ColHsn.Index].Value;
                        oSheet.Cells[xRow, Gst] = dgRow.Cells[ColGst.Index].Value;
                        oSheet.Cells[xRow, Pcs] = dgRow.Cells[ColPcs.Index].Value;
                        oSheet.Cells[xRow, SPrice] = dgRow.Cells[ColSalePrice.Index].Value;
                        oSheet.Cells[xRow, Rate] = dgRow.Cells[ColRate.Index].Value;
                        oSheet.Cells[xRow, Per] = "";
                        oSheet.Cells[xRow, Discount] = isNull(dgRow.Cells[ColDiscountP.Index].Value) ? 0 : dgRow.Cells[ColDiscountP.Index].Value.ToString() + "%";
                        
                        oSheet.Cells[xRow, Amount] = dgRow.Cells[ColAmount.Index].Value;

                        if (true)
                        {
                        Sgst += float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 200;       
                        }
                        else
                        {
                            Sgst += float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 100;
                        }

                        xRow += 1;

                    }
                }
                
                //SGST OR IGST
                if (true)
                {
                    xRow += 1;
                    oSheet.Cells[xRow, Description] = "SGST";
                    oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    oSheet.Cells[xRow, Amount] = Sgst;
                    xRow += 1;
                    oSheet.Cells[xRow, Description] = "CGST";
                    oSheet.Cells[xRow].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    oSheet.Cells[xRow, Amount] = Sgst;
                } else
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
                oSheet.Cells[xRow, Amount] = "=SUM("+"J" + dataFillStartRow.ToString()+":" +"J" + (xRow - 1).ToString()+")";
                //oSheet.Cells[xRow, Amount].NumberFormat = "0.00";
                //Total PCs
                oSheet.Cells[xRow, Pcs]= "=SUM(" + "E" + dataFillStartRow.ToString() + ":" + "E" + (xRow - 1).ToString() + ")";
                oSheet.get_Range("A"+xRow.ToString(), "J"+xRow.ToString()).Font.Bold = true;
                //

                //Formating Above data
                xRow += 1;
                oSheet.Cells[xRow, Srno] = "Amount Chargeable (in words)";
                oSheet.get_Range("A" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                
                xRow += 1;
                oSheet.Cells[xRow, Srno] = "Indian Rupees Twenty Three Thousand Only";
                oSheet.get_Range("A" + xRow.ToString(), "J" + xRow.ToString()).Merge();
                oSheet.get_Range("A" + (xRow-1).ToString(), "J" + xRow.ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;

                xRow += 1;
                oSheet.Cells[xRow, Srno] = "HSN/SAC";
                oSheet.get_Range("A" + xRow.ToString(), "C" + (xRow+2).ToString()).Merge();
                oSheet.get_Range("A" + (xRow).ToString(), "J" + (xRow + 2).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;

                oSheet.Cells[xRow, Gst] = "Taxable Value";
                oSheet.get_Range("D" + xRow.ToString(), "D" + (xRow + 2).ToString()).Merge();
                oSheet.get_Range("D" + (xRow).ToString(), "D" + (xRow + 2).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;

                //Central tax
                if (true)
                {
                    oSheet.Cells[xRow, Pcs] = "Central Tax";
                    oSheet.get_Range("E" + xRow.ToString(), "F" + (xRow).ToString()).Merge();
                    oSheet.get_Range("E" + (xRow).ToString(), "E" + (xRow).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    
                    oSheet.Cells[xRow+1, Pcs] = "Rate";
                    oSheet.Cells[xRow+1, SPrice] = "Amount";
                    oSheet.get_Range("E" + (xRow+1).ToString(), "E" + (xRow+2).ToString()).Merge();
                    oSheet.get_Range("F" + (xRow + 1).ToString(), "F" + (xRow + 2).ToString()).Merge();


                    oSheet.Cells[xRow, Rate] = "State Tax";
                    oSheet.get_Range("G" + xRow.ToString(), "I" + (xRow).ToString()).Merge();
                    oSheet.get_Range("G" + (xRow).ToString(), "I" + (xRow).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;

                    oSheet.Cells[xRow + 1, Rate] = "Rate";
                    oSheet.Cells[xRow + 1, Per] = "Amount";
                    oSheet.get_Range("G" + (xRow+1).ToString(), "G" + (xRow+2).ToString()).Merge();
                    oSheet.get_Range("H" + (xRow + 1).ToString(), "I" + (xRow + 2).ToString()).Merge();

                }
                else
                {
                    oSheet.Cells[xRow, Pcs] = "Inter State Tax";
                    oSheet.get_Range("E" + xRow.ToString(), "I" + (xRow).ToString()).Merge();
                    oSheet.get_Range("E" + (xRow).ToString(), "I" + (xRow).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;
                }

                //Total Tax
                oSheet.Cells[xRow, Amount] = "Total Tax Amount";
                oSheet.get_Range("J" + xRow.ToString(), "J" + (xRow+2).ToString()).Merge();
                oSheet.get_Range("J" + (xRow).ToString(), "J" + (xRow+2).ToString()).Cells.Borders.Weight = Excel.XlBorderWeight.xlThin;

                //Tax Calculation
                        xRow += 3;
                int taxStartRow = xRow;
                foreach (DataGridViewRow dgRow in DgvMain.Rows)
                {
                    if (dgRow.Index > 0)
                    {
                        oSheet.Cells[xRow, Srno] = dgRow.Cells[ColHsn.Index].Value;
                        oSheet.get_Range("A" + xRow.ToString(), "C" + xRow .ToString()).Merge();

                        oSheet.Cells[xRow, Gst] = dgRow.Cells[ColAmount.Index].Value;
                        if (true)
                        {
                            //central
                            oSheet.Cells[xRow, Pcs] = float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 2;
                            oSheet.Cells[xRow, SPrice] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 200;
                            //state
                            oSheet.Cells[xRow, SPrice] = float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 2;
                            oSheet.Cells[xRow, Per] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 200;

                            //total
                            oSheet.Cells[xRow, Amount] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 100;

                        }
                        else
                        {
                            oSheet.Cells[xRow, Pcs] =dgRow.Cells[ColGst.Index].Value;
                            oSheet.Cells[xRow, SPrice] = float.Parse((string)dgRow.Cells[ColAmount.Index].Value) * float.Parse((string)dgRow.Cells[ColGst.Index].Value) / 100;
                        }


                        xRow += 1;
                    }
                }

                //Tax Total 
                oSheet.Cells[xRow, Srno] = "Total";
                oSheet.get_Range("A" + xRow.ToString(), "C" + xRow.ToString()).Merge();

                
                oSheet.Cells[xRow, Gst] = "=SUM(" + "D" + taxStartRow.ToString() + ":" + "D" + (xRow - 1).ToString() + ")";

                if (true)
                {
                    oSheet.Cells[xRow, SPrice] = "=SUM(" + "F" + taxStartRow.ToString() + ":" + "F" + (xRow - 1).ToString() + ")";
                    oSheet.Cells[xRow, Pcs] = "=SUM(" + "H" + taxStartRow.ToString() + ":" + "H" + (xRow - 1).ToString() + ")";

                } else
                {
                    oSheet.Cells[xRow, Rate] = "=SUM(" + "H" + taxStartRow.ToString() + ":" + "H" + (xRow - 1).ToString() + ")";

                }

                oSheet.Cells[xRow, Amount] = "=SUM(" + "J" + taxStartRow.ToString() + ":" + "J" + (xRow - 1).ToString() + ")";

                //Disclamer


                #endregion

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "J30");
                oRng.EntireColumn.AutoFit();


                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
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

                    if (DialogResult.Yes == MessageBox.Show("Sure You Want To Delete This Row??", "Row Delete", MessageBoxButtons.YesNo,MessageBoxIcon.Question))
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
            } else
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
    }
}