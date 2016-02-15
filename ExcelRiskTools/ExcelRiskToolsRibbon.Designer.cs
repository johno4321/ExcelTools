namespace ExcelRiskTools
{
    partial class ExcelRiskToolsRibbon
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ExcelRiskToolsTab = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.DorisImportPrepGroup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.PrepSheetButton = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.ExcelRiskToolsTab.SuspendLayout();
            this.DorisImportPrepGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // ExcelRiskToolsTab
            // 
            this.ExcelRiskToolsTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ExcelRiskToolsTab.Groups.Add(this.DorisImportPrepGroup);
            this.ExcelRiskToolsTab.Label = "Excel Risk Tools";
            this.ExcelRiskToolsTab.Name = "ExcelRiskToolsTab";
            // 
            // DorisImportPrepGroup
            // 
            this.DorisImportPrepGroup.Items.Add(this.PrepSheetButton);
            this.DorisImportPrepGroup.Label = "Doris Import Prep";
            this.DorisImportPrepGroup.Name = "DorisImportPrepGroup";
            // 
            // PrepSheetButton
            // 
            this.PrepSheetButton.Image = global::ExcelRiskTools.Properties.Resources.Brush;
            this.PrepSheetButton.Label = "Prep Sheet";
            this.PrepSheetButton.Name = "PrepSheetButton";
            this.PrepSheetButton.ShowImage = true;
            this.PrepSheetButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.PrepSheetButton_Click);
            // 
            // ExcelRiskToolsRibbon
            // 
            this.Name = "ExcelRiskToolsRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ExcelRiskToolsTab);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.Ribbon1_Load);
            this.ExcelRiskToolsTab.ResumeLayout(false);
            this.ExcelRiskToolsTab.PerformLayout();
            this.DorisImportPrepGroup.ResumeLayout(false);
            this.DorisImportPrepGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ExcelRiskToolsTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DorisImportPrepGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PrepSheetButton;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal ExcelRiskToolsRibbon Ribbon1
        {
            get { return this.GetRibbon<ExcelRiskToolsRibbon>(); }
        }
    }
}
