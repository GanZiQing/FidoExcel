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
            this.FidoRibbon = this.Factory.CreateRibbonTab();
            this.AutomationToolsGroup = this.Factory.CreateRibbonGroup();
            this.ETABSPaneLauncher = this.Factory.CreateRibbonButton();
            this.reportPane = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ExcelToolsButton = this.Factory.CreateRibbonButton();
            this.FormatToolsButton = this.Factory.CreateRibbonButton();
            this.PlottingTools = this.Factory.CreateRibbonButton();
            this.forSharing = this.Factory.CreateRibbonGroup();
            this.DirectoryAndPdfButton = this.Factory.CreateRibbonButton();
            this.draftingPaneButton = this.Factory.CreateRibbonButton();
            this.toHide = this.Factory.CreateRibbonGroup();
            this.PilingToolsButton = this.Factory.CreateRibbonButton();
            this.beamDesign = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.wallDesign = this.Factory.CreateRibbonButton();
            this.wallCheck = this.Factory.CreateRibbonButton();
            this.Misc = this.Factory.CreateRibbonGroup();
            this.versionLabel = this.Factory.CreateRibbonLabel();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.FidoRibbon.SuspendLayout();
            this.AutomationToolsGroup.SuspendLayout();
            this.group2.SuspendLayout();
            this.forSharing.SuspendLayout();
            this.toHide.SuspendLayout();
            this.group1.SuspendLayout();
            this.Misc.SuspendLayout();
            this.tab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // FidoRibbon
            // 
            this.FidoRibbon.Groups.Add(this.AutomationToolsGroup);
            this.FidoRibbon.Groups.Add(this.group2);
            this.FidoRibbon.Groups.Add(this.forSharing);
            this.FidoRibbon.Groups.Add(this.toHide);
            this.FidoRibbon.Groups.Add(this.group1);
            this.FidoRibbon.Groups.Add(this.Misc);
            this.FidoRibbon.KeyTip = "L1";
            this.FidoRibbon.Label = "Fido";
            this.FidoRibbon.Name = "FidoRibbon";
            // 
            // AutomationToolsGroup
            // 
            this.AutomationToolsGroup.Items.Add(this.ETABSPaneLauncher);
            this.AutomationToolsGroup.Items.Add(this.reportPane);
            this.AutomationToolsGroup.Label = "ETABS";
            this.AutomationToolsGroup.Name = "AutomationToolsGroup";
            // 
            // ETABSPaneLauncher
            // 
            this.ETABSPaneLauncher.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ETABSPaneLauncher.Image = global::ExcelAddIn2.Properties.Resources.etabs;
            this.ETABSPaneLauncher.Label = "ETABS";
            this.ETABSPaneLauncher.Name = "ETABSPaneLauncher";
            this.ETABSPaneLauncher.ShowImage = true;
            this.ETABSPaneLauncher.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ETABSPaneLauncher_Click);
            // 
            // reportPane
            // 
            this.reportPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.reportPane.Image = global::ExcelAddIn2.Properties.Resources.ppt;
            this.reportPane.Label = "Report";
            this.reportPane.Name = "reportPane";
            this.reportPane.ShowImage = true;
            this.reportPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.reportPane_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.ExcelToolsButton);
            this.group2.Items.Add(this.FormatToolsButton);
            this.group2.Items.Add(this.PlottingTools);
            this.group2.Label = "Excel";
            this.group2.Name = "group2";
            // 
            // ExcelToolsButton
            // 
            this.ExcelToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExcelToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.ExcelToolsButton.Label = "Iteration";
            this.ExcelToolsButton.Name = "ExcelToolsButton";
            this.ExcelToolsButton.ShowImage = true;
            this.ExcelToolsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.IterationTools_Click);
            // 
            // FormatToolsButton
            // 
            this.FormatToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.FormatToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.FormatToolsButton.Label = "Formatting";
            this.FormatToolsButton.Name = "FormatToolsButton";
            this.FormatToolsButton.ShowImage = true;
            this.FormatToolsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestButt_Click);
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
            this.forSharing.Items.Add(this.DirectoryAndPdfButton);
            this.forSharing.Items.Add(this.draftingPaneButton);
            this.forSharing.Label = "Directories and PDF";
            this.forSharing.Name = "forSharing";
            // 
            // DirectoryAndPdfButton
            // 
            this.DirectoryAndPdfButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.DirectoryAndPdfButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.DirectoryAndPdfButton.Label = "Directory and PDF";
            this.DirectoryAndPdfButton.Name = "DirectoryAndPdfButton";
            this.DirectoryAndPdfButton.ShowImage = true;
            this.DirectoryAndPdfButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PrintToolsButton_Click);
            // 
            // draftingPaneButton
            // 
            this.draftingPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.draftingPaneButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.draftingPaneButton.Label = "Drafting";
            this.draftingPaneButton.Name = "draftingPaneButton";
            this.draftingPaneButton.ShowImage = true;
            this.draftingPaneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.draftingPaneButton_Click);
            // 
            // toHide
            // 
            this.toHide.Items.Add(this.PilingToolsButton);
            this.toHide.Items.Add(this.beamDesign);
            this.toHide.Label = "Hidden Group";
            this.toHide.Name = "toHide";
            this.toHide.Visible = false;
            // 
            // PilingToolsButton
            // 
            this.PilingToolsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.PilingToolsButton.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.PilingToolsButton.Label = "Piling Tools";
            this.PilingToolsButton.Name = "PilingToolsButton";
            this.PilingToolsButton.ShowImage = true;
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
            // group1
            // 
            this.group1.Items.Add(this.wallDesign);
            this.group1.Items.Add(this.wallCheck);
            this.group1.Label = "Wall Design";
            this.group1.Name = "group1";
            // 
            // wallDesign
            // 
            this.wallDesign.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.wallDesign.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.wallDesign.Label = "Wall Design";
            this.wallDesign.Name = "wallDesign";
            this.wallDesign.ShowImage = true;
            this.wallDesign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wallDesign_Click);
            // 
            // wallCheck
            // 
            this.wallCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.wallCheck.Image = global::ExcelAddIn2.Properties.Resources.excel;
            this.wallCheck.Label = "Wall Check";
            this.wallCheck.Name = "wallCheck";
            this.wallCheck.ShowImage = true;
            this.wallCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wallCheck_Click);
            // 
            // Misc
            // 
            this.Misc.Items.Add(this.versionLabel);
            this.Misc.Label = "Info";
            this.Misc.Name = "Misc";
            // 
            // versionLabel
            // 
            this.versionLabel.Label = "Version Info";
            this.versionLabel.Name = "versionLabel";
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
            this.Tabs.Add(this.FidoRibbon);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Launcher_Load);
            this.FidoRibbon.ResumeLayout(false);
            this.FidoRibbon.PerformLayout();
            this.AutomationToolsGroup.ResumeLayout(false);
            this.AutomationToolsGroup.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.forSharing.ResumeLayout(false);
            this.forSharing.PerformLayout();
            this.toHide.ResumeLayout(false);
            this.toHide.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Misc.ResumeLayout(false);
            this.Misc.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonTab FidoRibbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AutomationToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ETABSPaneLauncher;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExcelToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FormatToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PilingToolsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DirectoryAndPdfButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup forSharing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton reportPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PlottingTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton beamDesign;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup toHide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wallDesign;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton draftingPaneButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton wallCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Misc;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel versionLabel;
    }

    partial class ThisRibbonCollection
    {
        internal LauncherRib Launcher
        {
            get { return this.GetRibbon<LauncherRib>(); }
        }
    }
}
