using System;
using System.Windows.Forms;

namespace Cell_Analyzer
{
    public partial class CellAnalyzerPane : UserControl
    {
        public CellAnalyzerPane()
        {
            InitializeComponent();
        }

        private void btnUnicode_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range rangeCell;
            rangeCell = Globals.ThisAddIn.Application.ActiveCell;

            string cellValue = "";

            if (null != rangeCell.Value)
            {
                cellValue = rangeCell.Value.ToString();
            }

            //Output the result
            txtResult.Text = CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(cellValue);
        }

    }

}
