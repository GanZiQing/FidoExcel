namespace ExcelAddIn2
{
    partial class LauncherRib : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public LauncherRib()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.LauncherRibbon = this.Factory.CreateRibbonTab();
            this.AutomationToolsGroup = this.Factory.CreateRibbonGroup();
            this.ETABSPaneLauncher = this.Factory.CreateRibbonButton();
            this.ExcelToolsButton = this.Factory.CreateRibbonButton();
            this.FormatToolsButton = this.Factory.CreateRibbonButton();
            this.PrintToolsButton = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.PilingToolsButton = this.Factory.CreateRibbonButton();
            this.PlottingTools = this.Factory.CreateRibbonButton();
            this.forSharing = this.Factory.CreateRibbonGroup();
            this.reportPane = this.Factory.CreateRibbonButton();
            this.beamDesign = this.Factory.CreateRibbonButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.LauncherRibbon.SuspendLayout();
            this.AutomationToolsGroup.SuspendLayout();
            this.group1.SuspendLayout();
            this.forSharing.SuspendLayout();
            this.tab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // LauncherRibbon
            // 
            this.LauncherRibbon.Groups.Add(this.AutomationToolsGroup);
            this.LauncherRibbon.Groups.Add(this.group1);
            this.LauncherRibbon.Groups.Add(this.forSharing);
            this.LauncherRibbon.KeyTip = "L1";
            this.LauncherRibbon.Label = "Launcher";
            this.LauncherRibbon.Name = "LauncherRibbon";
            // 
            // AutomationToolsGroup
            // 
            this.AutomationToolsGroup.Items.Add(this.ETABSPaneLauncher);
            this.AutomationToolsGroup.Items.Add(this.ExcelToolsButton);
            this.AutomationToolsGroup.Items.Add(this.FormatToolsButton);
            this.AutomationToolsGroup.Items.Add(this.PrintToolsButton);
            this.AutomationToolsGroup.Label = "Automation Tools";
            this.AutomationToolsGroup.Name = "AutomationToolsGroup";
            // 
            // ETABSPaneLauncher
            // 
            this.ETABSPaneLauncher.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ETABSPaneLauncher.Image = global::ExcelAddIn2.Properties.Resources.etabs;
            this.ETABSPaneLauncher.Label = "ETABS Tools";
            this.ETABSPaneLauncher.Name = "ETABSPaneLauncher";
            this.ETABSPaneLauncher.ShowImage = true;
            this.ETABSPaneLauncher.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ETABSPaneLauncher_Click);
            // 
            // ExcelToolsButton
            // 
            this.ExcelToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExcelToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.ExcelToolsButton.Label = "Iteration Tools";
            this.ExcelToolsButton.Name = "ExcelToolsButton";
            this.ExcelToolsButton.ShowImage = true;
            this.ExcelToolsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IterationTools_Click);
            // 
            // FormatToolsButton
            // 
            this.FormatToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FormatToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.FormatToolsButton.Label = "Format Tools";
            this.FormatToolsButton.Name = "FormatToolsButton";
            this.FormatToolsButton.ShowImage = true;
            this.FormatToolsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestButt_Click);
            // 
            // PrintToolsButton
            // 
            this.PrintToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PrintToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.PrintToolsButton.Label = "Print Tools";
            this.PrintToolsButton.Name = "PrintToolsButton";
            this.PrintToolsButton.ShowImage = true;
            this.PrintToolsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PrintToolsButton_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.PilingToolsButton);
            this.group1.Items.Add(this.PlottingTools);
            this.group1.Label = "Geotech";
            this.group1.Name = "group1";
            // 
            // PilingToolsButton
            // 
            this.PilingToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PilingToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.PilingToolsButton.Label = "Piling Tools";
            this.PilingToolsButton.Name = "PilingToolsButton";
            this.PilingToolsButton.ShowImage = true;
            // 
            // PlottingTools
            // 
            this.PlottingTools.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PlottingTools.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.PlottingTools.Label = "Plotting Tools";
            this.PlottingTools.Name = "PlottingTools";
            this.PlottingTools.ShowImage = true;
            this.PlottingTools.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PlottingTools_Click);
            // 
            // forSharing
            // 
            this.forSharing.Items.Add(this.reportPane);
            this.forSharing.Items.Add(this.beamDesign);
            this.forSharing.Label = "For Sharing";
            this.forSharing.Name = "forSharing";
            // 
            // reportPane
            // 
            this.reportPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.reportPane.Image = global::ExcelAddIn2.Properties.Resources.ppt;
            this.reportPane.Label = "Report Pane";
            this.reportPane.Name = "reportPane";
            this.reportPane.ShowImage = true;
            this.reportPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.reportPane_Click);
            // 
            // beamDesign
            // 
            this.beamDesign.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.beamDesign.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.beamDesign.Label = "Beam Design";
            this.beamDesign.Name = "beamDesign";
            this.beamDesign.ShowImage = true;
            this.beamDesign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.beamDesign_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // LauncherRib
            // 
            this.Name = "LauncherRib";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.LauncherRibbon);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Launcher_Load);
            this.LauncherRibbon.ResumeLayout(false);
            this.LauncherRibbon.PerformLayout();
            this.AutomationToolsGroup.ResumeLayout(false);
            this.AutomationToolsGroup.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.forSharing.ResumeLayout(false);
            this.forSharing.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonTab LauncherRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AutomationToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ETABSPaneLauncher;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExcelToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FormatToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PilingToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PrintToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup forSharing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton reportPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PlottingTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton beamDesign;
    }

    partial class ThisRibbonCollection
    {
        internal LauncherRib Launcher
        {
            get { return this.GetRibbon<LauncherRib>(); }
        }
    }
}
