using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelAddIn2
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.myFirstTab = this.Factory.CreateRibbonTab();
            this.MagicTab = this.Factory.CreateRibbonTab();
            this.UnitDuplicator = this.Factory.CreateRibbonGroup();
            this.GetGroups = this.Factory.CreateRibbonButton();
            this.getNodeCoord = this.Factory.CreateRibbonButton();
            this.DuplicateUnits = this.Factory.CreateRibbonButton();
            this.GetFloors = this.Factory.CreateRibbonButton();
            this.CopyFrameLabel = this.Factory.CreateRibbonButton();
            this.RemoveUNBack = this.Factory.CreateRibbonButton();
            this.Utilities = this.Factory.CreateRibbonGroup();
            this.checkWalls = this.Factory.CreateRibbonButton();
            this.SelectBeamLabel = this.Factory.CreateRibbonButton();
            this.drawDropPanel = this.Factory.CreateRibbonButton();
            this.errorJoints = this.Factory.CreateRibbonButton();
            this.offsetter = this.Factory.CreateRibbonButton();
            this.GetJointFromFrame = this.Factory.CreateRibbonButton();
            this.setJCoord = this.Factory.CreateRibbonButton();
            this.SetOuputLC = this.Factory.CreateRibbonButton();
            this.GetPierForces = this.Factory.CreateRibbonButton();
            this.Report = this.Factory.CreateRibbonGroup();
            this.GetFiles = this.Factory.CreateRibbonButton();
            this.TestGroup = this.Factory.CreateRibbonGroup();
            this.AutocadTest = this.Factory.CreateRibbonButton();
            this.GetPierTest = this.Factory.CreateRibbonButton();
            this.Test = this.Factory.CreateRibbonButton();
            this.OutputJointLoad = this.Factory.CreateRibbonButton();
            this.GetJointTest = this.Factory.CreateRibbonButton();
            this.myFirstTab.SuspendLayout();
            this.MagicTab.SuspendLayout();
            this.UnitDuplicator.SuspendLayout();
            this.Utilities.SuspendLayout();
            this.Report.SuspendLayout();
            this.TestGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // myFirstTab
            // 
            this.myFirstTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.myFirstTab.Label = "TabAddIns";
            this.myFirstTab.Name = "myFirstTab";
            // 
            // MagicTab
            // 
            this.MagicTab.Groups.Add(this.UnitDuplicator);
            this.MagicTab.Groups.Add(this.Utilities);
            this.MagicTab.Groups.Add(this.Report);
            this.MagicTab.Groups.Add(this.TestGroup);
            this.MagicTab.Label = "Magic";
            this.MagicTab.Name = "MagicTab";
            // 
            // UnitDuplicator
            // 
            this.UnitDuplicator.Items.Add(this.GetGroups);
            this.UnitDuplicator.Items.Add(this.getNodeCoord);
            this.UnitDuplicator.Items.Add(this.DuplicateUnits);
            this.UnitDuplicator.Items.Add(this.GetFloors);
            this.UnitDuplicator.Items.Add(this.CopyFrameLabel);
            this.UnitDuplicator.Items.Add(this.RemoveUNBack);
            this.UnitDuplicator.Label = "Unit Duplicator";
            this.UnitDuplicator.Name = "UnitDuplicator";
            // 
            // GetGroups
            // 
            this.GetGroups.Label = "Get Groups";
            this.GetGroups.Name = "GetGroups";
            this.GetGroups.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetGroups_Click);
            // 
            // getNodeCoord
            // 
            this.getNodeCoord.Label = "Get Selected Coord";
            this.getNodeCoord.Name = "getNodeCoord";
            this.getNodeCoord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.getNodeCoord_Click);
            // 
            // DuplicateUnits
            // 
            this.DuplicateUnits.Label = "Duplicate Units";
            this.DuplicateUnits.Name = "DuplicateUnits";
            this.DuplicateUnits.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DuplicateUnits_Click);
            // 
            // GetFloors
            // 
            this.GetFloors.Label = "Get Floors";
            this.GetFloors.Name = "GetFloors";
            this.GetFloors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetFloors_Click);
            // 
            // CopyFrameLabel
            // 
            this.CopyFrameLabel.Label = "Copy Frame Label";
            this.CopyFrameLabel.Name = "CopyFrameLabel";
            this.CopyFrameLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CopyFrameLabel_Click);
            // 
            // RemoveUNBack
            // 
            this.RemoveUNBack.Label = "RemoveUNBack";
            this.RemoveUNBack.Name = "RemoveUNBack";
            this.RemoveUNBack.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveUNBack_Click);
            // 
            // Utilities
            // 
            this.Utilities.Items.Add(this.checkWalls);
            this.Utilities.Items.Add(this.SelectBeamLabel);
            this.Utilities.Items.Add(this.drawDropPanel);
            this.Utilities.Items.Add(this.errorJoints);
            this.Utilities.Items.Add(this.offsetter);
            this.Utilities.Items.Add(this.GetJointFromFrame);
            this.Utilities.Items.Add(this.setJCoord);
            this.Utilities.Items.Add(this.SetOuputLC);
            this.Utilities.Items.Add(this.GetPierForces);
            this.Utilities.Label = "Utilities";
            this.Utilities.Name = "Utilities";
            // 
            // checkWalls
            // 
            this.checkWalls.Label = "checkWalls";
            this.checkWalls.Name = "checkWalls";
            this.checkWalls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkWalls_Click);
            // 
            // SelectBeamLabel
            // 
            this.SelectBeamLabel.Label = "SelectBeamLabel";
            this.SelectBeamLabel.Name = "SelectBeamLabel";
            this.SelectBeamLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectBeamLabel_Click);
            // 
            // drawDropPanel
            // 
            this.drawDropPanel.Label = "Draw Drop Panel";
            this.drawDropPanel.Name = "drawDropPanel";
            this.drawDropPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drawDropPanel_Click);
            // 
            // errorJoints
            // 
            this.errorJoints.Label = "Error Joints";
            this.errorJoints.Name = "errorJoints";
            this.errorJoints.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.errorJoints_Click);
            // 
            // offsetter
            // 
            this.offsetter.Label = "Offsetter";
            this.offsetter.Name = "offsetter";
            this.offsetter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.offsetter_Click);
            // 
            // GetJointFromFrame
            // 
            this.GetJointFromFrame.Label = "Get J Fr F";
            this.GetJointFromFrame.Name = "GetJointFromFrame";
            this.GetJointFromFrame.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetJointFromFrame_Click);
            // 
            // setJCoord
            // 
            this.setJCoord.Label = "Set J Coord";
            this.setJCoord.Name = "setJCoord";
            this.setJCoord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.setJCoord_Click);
            // 
            // SetOuputLC
            // 
            this.SetOuputLC.Label = "SetOuputLC";
            this.SetOuputLC.Name = "SetOuputLC";
            this.SetOuputLC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetOuputLC_Click);
            // 
            // GetPierForces
            // 
            this.GetPierForces.Label = "GetPierForces";
            this.GetPierForces.Name = "GetPierForces";
            this.GetPierForces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetPierForces_Click);
            // 
            // Report
            // 
            this.Report.Items.Add(this.GetFiles);
            this.Report.Label = "Report";
            this.Report.Name = "Report";
            // 
            // GetFiles
            // 
            this.GetFiles.Label = "GetFiles";
            this.GetFiles.Name = "GetFiles";
            this.GetFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetFiles_Click);
            // 
            // TestGroup
            // 
            this.TestGroup.Items.Add(this.AutocadTest);
            this.TestGroup.Items.Add(this.GetPierTest);
            this.TestGroup.Items.Add(this.Test);
            this.TestGroup.Items.Add(this.OutputJointLoad);
            this.TestGroup.Items.Add(this.GetJointTest);
            this.TestGroup.Label = "Test Group";
            this.TestGroup.Name = "TestGroup";
            // 
            // AutocadTest
            // 
            this.AutocadTest.Label = "AutocadTest";
            this.AutocadTest.Name = "AutocadTest";
            this.AutocadTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AutocadTest_Click);
            // 
            // GetPierTest
            // 
            this.GetPierTest.Label = "GetPierTest";
            this.GetPierTest.Name = "GetPierTest";
            this.GetPierTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetPierTest_Click);
            // 
            // Test
            // 
            this.Test.Label = "TestSaveUniqueData";
            this.Test.Name = "Test";
            this.Test.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Test_Click);
            // 
            // OutputJointLoad
            // 
            this.OutputJointLoad.Label = "OutputJointLoad";
            this.OutputJointLoad.Name = "OutputJointLoad";
            this.OutputJointLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetJointLoad_Click);
            // 
            // GetJointTest
            // 
            this.GetJointTest.Label = "GetJointTest";
            this.GetJointTest.Name = "GetJointTest";
            this.GetJointTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetJointTest_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.myFirstTab);
            this.Tabs.Add(this.MagicTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.myFirstTab.ResumeLayout(false);
            this.myFirstTab.PerformLayout();
            this.MagicTab.ResumeLayout(false);
            this.MagicTab.PerformLayout();
            this.UnitDuplicator.ResumeLayout(false);
            this.UnitDuplicator.PerformLayout();
            this.Utilities.ResumeLayout(false);
            this.Utilities.PerformLayout();
            this.Report.ResumeLayout(false);
            this.Report.PerformLayout();
            this.TestGroup.ResumeLayout(false);
            this.TestGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab myFirstTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab MagicTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup UnitDuplicator;
        internal RibbonButton GetGroups;
        internal RibbonButton DuplicateUnits;
        internal RibbonButton OutputJointLoad;
        internal RibbonButton GetFloors;
        internal RibbonGroup TestGroup;
        internal RibbonButton CopyFrameLabel;
        internal RibbonButton RemoveUNBack;
        internal RibbonGroup Utilities;
        internal RibbonButton SelectBeamLabel;
        internal RibbonButton checkWalls;
        internal RibbonButton drawDropPanel;
        internal RibbonButton getNodeCoord;
        internal RibbonButton errorJoints;
        internal RibbonButton offsetter;
        internal RibbonButton GetJointFromFrame;
        internal RibbonButton setJCoord;
        internal RibbonButton GetPierForces;
        internal RibbonButton SetOuputLC;
        internal RibbonGroup Report;
        internal RibbonButton GetFiles;
        internal RibbonButton Test;
        internal RibbonButton GetPierTest;
        internal RibbonButton GetJointTest;
        internal RibbonButton AutocadTest;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }

}
