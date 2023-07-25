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
        Dictionary<string, string> states = new Dictionary<string, string>();
        private void Home_Load(object sender, EventArgs e)
        {
            //Adding Value in Dictionary
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
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";
                oSheet.Cells[1, 5] = "Tax";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
                DisplayQuarterlySales(oSheet);

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

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            String sMsg;
            int iNumQtrs;

            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

            MessageBox.Show(sMsg, "Quarterly Sales");

            //Starting at E1, fill headers for the number of columns selected.
            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            oResizeRange.Interior.ColorIndex = 36;

            //Fill the columns with a formula and apply a number format.
            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            //Apply borders to the Sales data and headers.
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Add a Totals formula for the sales data and apply a border.
            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
            = Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
            = Excel.XlBorderWeight.xlThick;

            //Add a Chart for the selected data.
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
            Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
            Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);

            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
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
    }
}