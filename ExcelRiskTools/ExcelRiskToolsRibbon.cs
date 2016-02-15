using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Tools.Ribbon;
using DocumentFormat.OpenXml;

namespace ExcelRiskTools
{
    public partial class ExcelRiskToolsRibbon : OfficeRibbon
    {
        public ExcelRiskToolsRibbon()
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void PrepSheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (activeWorkbook == null || !activeWorkbook.AreThereAnySheets())
            {
                MessageBox.Show(@"no workbook - see ya!");
            }

            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            var activeSheetName = activeSheet.Name;
            
            var spreadsheetDocument = SpreadsheetDocument.Open(activeWorkbook.FullName, false);
        }
    }
}
