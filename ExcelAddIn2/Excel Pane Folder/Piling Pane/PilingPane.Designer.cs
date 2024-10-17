namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class PilingPane
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
            this.pilingTabControl = new System.Windows.Forms.TabControl();
            this.agsTab = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.removeCont = new System.Windows.Forms.Button();
            this.agsGroup = new System.Windows.Forms.GroupBox();
            this.checkCompressOutput = new System.Windows.Forms.CheckBox();
            this.dispSpt100Range = new System.Windows.Forms.TextBox();
            this.setSpt100Range = new System.Windows.Forms.Button();
            this.checkDefaultSPT = new System.Windows.Forms.CheckBox();
            this.basicAGS = new System.Windows.Forms.Button();
            this.DispRockRange = new System.Windows.Forms.TextBox();
            this.printDescriptionCheck = new System.Windows.Forms.CheckBox();
            this.SetRockRange = new System.Windows.Forms.Button();
            this.checkFillSoilType = new System.Windows.Forms.CheckBox();
            this.checkRemoveNoSPT = new System.Windows.Forms.CheckBox();
            this.importAGS = new System.Windows.Forms.Button();
            this.bhGroup = new System.Windows.Forms.GroupBox();
            this.checkDrawRef = new System.Windows.Forms.CheckBox();
            this.DispNotRockRange = new System.Windows.Forms.TextBox();
            this.SetNotRockRange = new System.Windows.Forms.Button();
            this.DrawBH = new System.Windows.Forms.Button();
            this.DispDrawRange = new System.Windows.Forms.TextBox();
            this.SetDrawRange = new System.Windows.Forms.Button();
            this.pilingDesignTab = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.delSheets = new System.Windows.Forms.Button();
            this.dispEfficiencyUpper = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dispEfficiencyLower = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.checkDeactivateScreen = new System.Windows.Forms.CheckBox();
            this.designPiles = new System.Windows.Forms.Button();
            this.setSheetsToRun = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.dispAppendName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dispNsfTypeInput = new System.Windows.Forms.TextBox();
            this.setNsfTypeInput = new System.Windows.Forms.Button();
            this.dispRockTypeInput = new System.Windows.Forms.TextBox();
            this.setRockTypeInput = new System.Windows.Forms.Button();
            this.copySoilData = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dispBhRlCell = new System.Windows.Forms.TextBox();
            this.setBhRlCell = new System.Windows.Forms.Button();
            this.dispRockStart = new System.Windows.Forms.TextBox();
            this.setRockStart = new System.Windows.Forms.Button();
            this.dispEffRange = new System.Windows.Forms.TextBox();
            this.setEffRange = new System.Windows.Forms.Button();
            this.dispQbRange = new System.Windows.Forms.TextBox();
            this.setQbRange = new System.Windows.Forms.Button();
            this.dispFsRange = new System.Windows.Forms.TextBox();
            this.setFsRange = new System.Windows.Forms.Button();
            this.dispSoilDest = new System.Windows.Forms.TextBox();
            this.setSoilDest = new System.Windows.Forms.Button();
            this.dispRefSheet = new System.Windows.Forms.TextBox();
            this.setRefSheet = new System.Windows.Forms.Button();
            this.dispSoilInputData = new System.Windows.Forms.TextBox();
            this.setSoilInputData = new System.Windows.Forms.Button();
            this.dispSpt100Start = new System.Windows.Forms.TextBox();
            this.setSpt100Start = new System.Windows.Forms.Button();
            this.pilingTabControl.SuspendLayout();
            this.agsTab.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.agsGroup.SuspendLayout();
            this.bhGroup.SuspendLayout();
            this.pilingDesignTab.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pilingTabControl
            // 
            this.pilingTabControl.Controls.Add(this.agsTab);
            this.pilingTabControl.Controls.Add(this.pilingDesignTab);
            this.pilingTabControl.Location = new System.Drawing.Point(3, 3);
            this.pilingTabControl.Name = "pilingTabControl";
            this.pilingTabControl.SelectedIndex = 0;
            this.pilingTabControl.Size = new System.Drawing.Size(294, 1060);
            this.pilingTabControl.TabIndex = 0;
            // 
            // agsTab
            // 
            this.agsTab.BackColor = System.Drawing.SystemColors.Control;
            this.agsTab.Controls.Add(this.groupBox3);
            this.agsTab.Controls.Add(this.agsGroup);
            this.agsTab.Controls.Add(this.bhGroup);
            this.agsTab.Location = new System.Drawing.Point(4, 22);
            this.agsTab.Name = "agsTab";
            this.agsTab.Padding = new System.Windows.Forms.Padding(3);
            this.agsTab.Size = new System.Drawing.Size(286, 1034);
            this.agsTab.TabIndex = 0;
            this.agsTab.Text = "AGS";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.removeCont);
            this.groupBox3.Location = new System.Drawing.Point(6, 6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(274, 54);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Preprocess AGS";
            // 
            // removeCont
            // 
            this.removeCont.ForeColor = System.Drawing.SystemColors.WindowText;
            this.removeCont.Location = new System.Drawing.Point(6, 19);
            this.removeCont.Name = "removeCont";
            this.removeCont.Size = new System.Drawing.Size(262, 24);
            this.removeCont.TabIndex = 17;
            this.removeCont.Text = "Remove <Cont>";
            this.removeCont.UseVisualStyleBackColor = true;
            this.removeCont.Click += new System.EventHandler(this.removeCont_Click);
            // 
            // agsGroup
            // 
            this.agsGroup.Controls.Add(this.checkCompressOutput);
            this.agsGroup.Controls.Add(this.dispSpt100Range);
            this.agsGroup.Controls.Add(this.setSpt100Range);
            this.agsGroup.Controls.Add(this.checkDefaultSPT);
            this.agsGroup.Controls.Add(this.basicAGS);
            this.agsGroup.Controls.Add(this.DispRockRange);
            this.agsGroup.Controls.Add(this.printDescriptionCheck);
            this.agsGroup.Controls.Add(this.SetRockRange);
            this.agsGroup.Controls.Add(this.checkFillSoilType);
            this.agsGroup.Controls.Add(this.checkRemoveNoSPT);
            this.agsGroup.Controls.Add(this.importAGS);
            this.agsGroup.Location = new System.Drawing.Point(6, 66);
            this.agsGroup.Name = "agsGroup";
            this.agsGroup.Size = new System.Drawing.Size(274, 204);
            this.agsGroup.TabIndex = 1;
            this.agsGroup.TabStop = false;
            this.agsGroup.Text = "AGS(SG) Reader";
            // 
            // checkCompressOutput
            // 
            this.checkCompressOutput.AutoSize = true;
            this.checkCompressOutput.Checked = true;
            this.checkCompressOutput.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkCompressOutput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkCompressOutput.Location = new System.Drawing.Point(139, 46);
            this.checkCompressOutput.Name = "checkCompressOutput";
            this.checkCompressOutput.Size = new System.Drawing.Size(130, 17);
            this.checkCompressOutput.TabIndex = 24;
            this.checkCompressOutput.Text = "Combine Similar Rows";
            this.checkCompressOutput.UseVisualStyleBackColor = true;
            // 
            // dispSpt100Range
            // 
            this.dispSpt100Range.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispSpt100Range.Location = new System.Drawing.Point(134, 126);
            this.dispSpt100Range.Name = "dispSpt100Range";
            this.dispSpt100Range.Size = new System.Drawing.Size(134, 20);
            this.dispSpt100Range.TabIndex = 23;
            this.dispSpt100Range.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispSpt100Range.WordWrap = false;
            // 
            // setSpt100Range
            // 
            this.setSpt100Range.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSpt100Range.Location = new System.Drawing.Point(6, 123);
            this.setSpt100Range.Name = "setSpt100Range";
            this.setSpt100Range.Size = new System.Drawing.Size(122, 25);
            this.setSpt100Range.TabIndex = 22;
            this.setSpt100Range.Text = "Set SPT100 Type";
            this.setSpt100Range.UseVisualStyleBackColor = true;
            // 
            // checkDefaultSPT
            // 
            this.checkDefaultSPT.AutoSize = true;
            this.checkDefaultSPT.Checked = true;
            this.checkDefaultSPT.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkDefaultSPT.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkDefaultSPT.Location = new System.Drawing.Point(6, 46);
            this.checkDefaultSPT.Name = "checkDefaultSPT";
            this.checkDefaultSPT.Size = new System.Drawing.Size(99, 17);
            this.checkDefaultSPT.TabIndex = 20;
            this.checkDefaultSPT.Text = "Fill Default SPT";
            this.checkDefaultSPT.UseVisualStyleBackColor = true;
            // 
            // basicAGS
            // 
            this.basicAGS.ForeColor = System.Drawing.SystemColors.WindowText;
            this.basicAGS.Location = new System.Drawing.Point(146, 169);
            this.basicAGS.Name = "basicAGS";
            this.basicAGS.Size = new System.Drawing.Size(122, 24);
            this.basicAGS.TabIndex = 18;
            this.basicAGS.Text = "Import Raw AGS";
            this.basicAGS.UseVisualStyleBackColor = true;
            this.basicAGS.Click += new System.EventHandler(this.basicAGS_Click);
            // 
            // DispRockRange
            // 
            this.DispRockRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DispRockRange.Location = new System.Drawing.Point(134, 95);
            this.DispRockRange.Name = "DispRockRange";
            this.DispRockRange.Size = new System.Drawing.Size(134, 20);
            this.DispRockRange.TabIndex = 15;
            this.DispRockRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.DispRockRange.WordWrap = false;
            // 
            // printDescriptionCheck
            // 
            this.printDescriptionCheck.AutoSize = true;
            this.printDescriptionCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printDescriptionCheck.Location = new System.Drawing.Point(6, 69);
            this.printDescriptionCheck.Name = "printDescriptionCheck";
            this.printDescriptionCheck.Size = new System.Drawing.Size(104, 17);
            this.printDescriptionCheck.TabIndex = 19;
            this.printDescriptionCheck.Text = "Add Description ";
            this.printDescriptionCheck.UseVisualStyleBackColor = true;
            // 
            // SetRockRange
            // 
            this.SetRockRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.SetRockRange.Location = new System.Drawing.Point(6, 92);
            this.SetRockRange.Name = "SetRockRange";
            this.SetRockRange.Size = new System.Drawing.Size(122, 25);
            this.SetRockRange.TabIndex = 14;
            this.SetRockRange.Text = "Set Rock Type";
            this.SetRockRange.UseVisualStyleBackColor = true;
            // 
            // checkFillSoilType
            // 
            this.checkFillSoilType.AutoSize = true;
            this.checkFillSoilType.Checked = true;
            this.checkFillSoilType.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkFillSoilType.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkFillSoilType.Location = new System.Drawing.Point(139, 23);
            this.checkFillSoilType.Name = "checkFillSoilType";
            this.checkFillSoilType.Size = new System.Drawing.Size(85, 17);
            this.checkFillSoilType.TabIndex = 18;
            this.checkFillSoilType.Text = "Fill Soil Type";
            this.checkFillSoilType.UseVisualStyleBackColor = true;
            // 
            // checkRemoveNoSPT
            // 
            this.checkRemoveNoSPT.AutoSize = true;
            this.checkRemoveNoSPT.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkRemoveNoSPT.Location = new System.Drawing.Point(6, 23);
            this.checkRemoveNoSPT.Name = "checkRemoveNoSPT";
            this.checkRemoveNoSPT.Size = new System.Drawing.Size(122, 17);
            this.checkRemoveNoSPT.TabIndex = 17;
            this.checkRemoveNoSPT.Text = "Remove Empty SPT";
            this.checkRemoveNoSPT.UseVisualStyleBackColor = true;
            // 
            // importAGS
            // 
            this.importAGS.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importAGS.Location = new System.Drawing.Point(6, 169);
            this.importAGS.Name = "importAGS";
            this.importAGS.Size = new System.Drawing.Size(122, 24);
            this.importAGS.TabIndex = 16;
            this.importAGS.Text = "Import AGS";
            this.importAGS.UseVisualStyleBackColor = true;
            this.importAGS.Click += new System.EventHandler(this.importAGS_Click);
            // 
            // bhGroup
            // 
            this.bhGroup.Controls.Add(this.checkDrawRef);
            this.bhGroup.Controls.Add(this.DispNotRockRange);
            this.bhGroup.Controls.Add(this.SetNotRockRange);
            this.bhGroup.Controls.Add(this.DrawBH);
            this.bhGroup.Controls.Add(this.DispDrawRange);
            this.bhGroup.Controls.Add(this.SetDrawRange);
            this.bhGroup.Location = new System.Drawing.Point(6, 276);
            this.bhGroup.Name = "bhGroup";
            this.bhGroup.Size = new System.Drawing.Size(274, 143);
            this.bhGroup.TabIndex = 0;
            this.bhGroup.TabStop = false;
            this.bhGroup.Text = "Draw Boreholes";
            // 
            // checkDrawRef
            // 
            this.checkDrawRef.AutoSize = true;
            this.checkDrawRef.Checked = true;
            this.checkDrawRef.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkDrawRef.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkDrawRef.Location = new System.Drawing.Point(6, 81);
            this.checkDrawRef.Name = "checkDrawRef";
            this.checkDrawRef.Size = new System.Drawing.Size(127, 17);
            this.checkDrawRef.TabIndex = 20;
            this.checkDrawRef.Text = "Draw Reference Line";
            this.checkDrawRef.UseVisualStyleBackColor = true;
            // 
            // DispNotRockRange
            // 
            this.DispNotRockRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DispNotRockRange.Location = new System.Drawing.Point(134, 53);
            this.DispNotRockRange.Name = "DispNotRockRange";
            this.DispNotRockRange.Size = new System.Drawing.Size(134, 20);
            this.DispNotRockRange.TabIndex = 17;
            this.DispNotRockRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.DispNotRockRange.WordWrap = false;
            // 
            // SetNotRockRange
            // 
            this.SetNotRockRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.SetNotRockRange.Location = new System.Drawing.Point(6, 50);
            this.SetNotRockRange.Name = "SetNotRockRange";
            this.SetNotRockRange.Size = new System.Drawing.Size(122, 25);
            this.SetNotRockRange.TabIndex = 16;
            this.SetNotRockRange.Text = "Set Overwrite Type";
            this.SetNotRockRange.UseVisualStyleBackColor = true;
            // 
            // DrawBH
            // 
            this.DrawBH.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DrawBH.Location = new System.Drawing.Point(6, 104);
            this.DrawBH.Name = "DrawBH";
            this.DrawBH.Size = new System.Drawing.Size(262, 24);
            this.DrawBH.TabIndex = 13;
            this.DrawBH.Text = "Run";
            this.DrawBH.UseVisualStyleBackColor = true;
            this.DrawBH.Click += new System.EventHandler(this.DrawBH_Click);
            // 
            // DispDrawRange
            // 
            this.DispDrawRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DispDrawRange.Location = new System.Drawing.Point(134, 22);
            this.DispDrawRange.Name = "DispDrawRange";
            this.DispDrawRange.Size = new System.Drawing.Size(134, 20);
            this.DispDrawRange.TabIndex = 12;
            this.DispDrawRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.DispDrawRange.WordWrap = false;
            // 
            // SetDrawRange
            // 
            this.SetDrawRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.SetDrawRange.Location = new System.Drawing.Point(6, 19);
            this.SetDrawRange.Name = "SetDrawRange";
            this.SetDrawRange.Size = new System.Drawing.Size(122, 25);
            this.SetDrawRange.TabIndex = 11;
            this.SetDrawRange.Text = "Set Draw Range";
            this.SetDrawRange.UseVisualStyleBackColor = true;
            // 
            // pilingDesignTab
            // 
            this.pilingDesignTab.BackColor = System.Drawing.SystemColors.Control;
            this.pilingDesignTab.Controls.Add(this.groupBox2);
            this.pilingDesignTab.Controls.Add(this.groupBox1);
            this.pilingDesignTab.Location = new System.Drawing.Point(4, 22);
            this.pilingDesignTab.Name = "pilingDesignTab";
            this.pilingDesignTab.Padding = new System.Windows.Forms.Padding(3);
            this.pilingDesignTab.Size = new System.Drawing.Size(286, 1034);
            this.pilingDesignTab.TabIndex = 1;
            this.pilingDesignTab.Text = "Piling Design";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.delSheets);
            this.groupBox2.Controls.Add(this.dispEfficiencyUpper);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.dispEfficiencyLower);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.checkDeactivateScreen);
            this.groupBox2.Controls.Add(this.designPiles);
            this.groupBox2.Controls.Add(this.setSheetsToRun);
            this.groupBox2.Location = new System.Drawing.Point(6, 543);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(274, 175);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Design Sheets";
            // 
            // delSheets
            // 
            this.delSheets.ForeColor = System.Drawing.SystemColors.WindowText;
            this.delSheets.Location = new System.Drawing.Point(146, 19);
            this.delSheets.Name = "delSheets";
            this.delSheets.Size = new System.Drawing.Size(122, 25);
            this.delSheets.TabIndex = 51;
            this.delSheets.Text = "Delete Sheets";
            this.delSheets.UseVisualStyleBackColor = true;
            // 
            // dispEfficiencyUpper
            // 
            this.dispEfficiencyUpper.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispEfficiencyUpper.Location = new System.Drawing.Point(134, 81);
            this.dispEfficiencyUpper.Name = "dispEfficiencyUpper";
            this.dispEfficiencyUpper.Size = new System.Drawing.Size(134, 20);
            this.dispEfficiencyUpper.TabIndex = 50;
            this.dispEfficiencyUpper.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispEfficiencyUpper.WordWrap = false;
            // 
            // label3
            // 
            this.label3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label3.Location = new System.Drawing.Point(6, 76);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(122, 25);
            this.label3.TabIndex = 49;
            this.label3.Text = "Efficiency Upper Bound";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dispEfficiencyLower
            // 
            this.dispEfficiencyLower.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispEfficiencyLower.Location = new System.Drawing.Point(134, 52);
            this.dispEfficiencyLower.Name = "dispEfficiencyLower";
            this.dispEfficiencyLower.Size = new System.Drawing.Size(134, 20);
            this.dispEfficiencyLower.TabIndex = 48;
            this.dispEfficiencyLower.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispEfficiencyLower.WordWrap = false;
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label2.Location = new System.Drawing.Point(6, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(122, 25);
            this.label2.TabIndex = 47;
            this.label2.Text = "Efficiency Lower Bound";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // checkDeactivateScreen
            // 
            this.checkDeactivateScreen.AutoSize = true;
            this.checkDeactivateScreen.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkDeactivateScreen.Location = new System.Drawing.Point(6, 150);
            this.checkDeactivateScreen.Name = "checkDeactivateScreen";
            this.checkDeactivateScreen.Size = new System.Drawing.Size(115, 17);
            this.checkDeactivateScreen.TabIndex = 46;
            this.checkDeactivateScreen.Text = "Deactivate Screen";
            this.checkDeactivateScreen.UseVisualStyleBackColor = true;
            // 
            // designPiles
            // 
            this.designPiles.ForeColor = System.Drawing.SystemColors.WindowText;
            this.designPiles.Location = new System.Drawing.Point(6, 119);
            this.designPiles.Name = "designPiles";
            this.designPiles.Size = new System.Drawing.Size(262, 25);
            this.designPiles.TabIndex = 45;
            this.designPiles.Text = "Design Selected Sheets";
            this.designPiles.UseVisualStyleBackColor = true;
            this.designPiles.Click += new System.EventHandler(this.designPiles_Click);
            // 
            // setSheetsToRun
            // 
            this.setSheetsToRun.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSheetsToRun.Location = new System.Drawing.Point(6, 19);
            this.setSheetsToRun.Name = "setSheetsToRun";
            this.setSheetsToRun.Size = new System.Drawing.Size(122, 25);
            this.setSheetsToRun.TabIndex = 15;
            this.setSheetsToRun.Text = "Set Sheets To Run";
            this.setSheetsToRun.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dispSpt100Start);
            this.groupBox1.Controls.Add(this.setSpt100Start);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.dispAppendName);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.dispNsfTypeInput);
            this.groupBox1.Controls.Add(this.setNsfTypeInput);
            this.groupBox1.Controls.Add(this.dispRockTypeInput);
            this.groupBox1.Controls.Add(this.setRockTypeInput);
            this.groupBox1.Controls.Add(this.copySoilData);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.dispBhRlCell);
            this.groupBox1.Controls.Add(this.setBhRlCell);
            this.groupBox1.Controls.Add(this.dispRockStart);
            this.groupBox1.Controls.Add(this.setRockStart);
            this.groupBox1.Controls.Add(this.dispEffRange);
            this.groupBox1.Controls.Add(this.setEffRange);
            this.groupBox1.Controls.Add(this.dispQbRange);
            this.groupBox1.Controls.Add(this.setQbRange);
            this.groupBox1.Controls.Add(this.dispFsRange);
            this.groupBox1.Controls.Add(this.setFsRange);
            this.groupBox1.Controls.Add(this.dispSoilDest);
            this.groupBox1.Controls.Add(this.setSoilDest);
            this.groupBox1.Controls.Add(this.dispRefSheet);
            this.groupBox1.Controls.Add(this.setRefSheet);
            this.groupBox1.Controls.Add(this.dispSoilInputData);
            this.groupBox1.Controls.Add(this.setSoilInputData);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(274, 531);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Copy to Spreadsheet";
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label6.Location = new System.Drawing.Point(3, 428);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(125, 25);
            this.label6.TabIndex = 38;
            this.label6.Text = "Append Sheet Name";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dispAppendName
            // 
            this.dispAppendName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAppendName.Location = new System.Drawing.Point(134, 431);
            this.dispAppendName.Name = "dispAppendName";
            this.dispAppendName.Size = new System.Drawing.Size(134, 20);
            this.dispAppendName.TabIndex = 37;
            this.dispAppendName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAppendName.WordWrap = false;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label5.Location = new System.Drawing.Point(6, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(262, 20);
            this.label5.TabIndex = 35;
            this.label5.Text = "Set Input Parameters:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label4.Location = new System.Drawing.Point(6, 402);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(262, 20);
            this.label4.TabIndex = 34;
            this.label4.Text = "Set Sheet Details:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // dispNsfTypeInput
            // 
            this.dispNsfTypeInput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispNsfTypeInput.Location = new System.Drawing.Point(134, 109);
            this.dispNsfTypeInput.Name = "dispNsfTypeInput";
            this.dispNsfTypeInput.Size = new System.Drawing.Size(134, 20);
            this.dispNsfTypeInput.TabIndex = 33;
            this.dispNsfTypeInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispNsfTypeInput.WordWrap = false;
            // 
            // setNsfTypeInput
            // 
            this.setNsfTypeInput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setNsfTypeInput.Location = new System.Drawing.Point(6, 106);
            this.setNsfTypeInput.Name = "setNsfTypeInput";
            this.setNsfTypeInput.Size = new System.Drawing.Size(122, 25);
            this.setNsfTypeInput.TabIndex = 32;
            this.setNsfTypeInput.Text = "Set NSF Type Range";
            this.setNsfTypeInput.UseVisualStyleBackColor = true;
            // 
            // dispRockTypeInput
            // 
            this.dispRockTypeInput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRockTypeInput.Location = new System.Drawing.Point(134, 78);
            this.dispRockTypeInput.Name = "dispRockTypeInput";
            this.dispRockTypeInput.Size = new System.Drawing.Size(134, 20);
            this.dispRockTypeInput.TabIndex = 31;
            this.dispRockTypeInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRockTypeInput.WordWrap = false;
            // 
            // setRockTypeInput
            // 
            this.setRockTypeInput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRockTypeInput.Location = new System.Drawing.Point(6, 75);
            this.setRockTypeInput.Name = "setRockTypeInput";
            this.setRockTypeInput.Size = new System.Drawing.Size(122, 25);
            this.setRockTypeInput.TabIndex = 30;
            this.setRockTypeInput.Text = "Set Rock Type Range";
            this.setRockTypeInput.UseVisualStyleBackColor = true;
            // 
            // copySoilData
            // 
            this.copySoilData.ForeColor = System.Drawing.SystemColors.WindowText;
            this.copySoilData.Location = new System.Drawing.Point(6, 467);
            this.copySoilData.Name = "copySoilData";
            this.copySoilData.Size = new System.Drawing.Size(262, 25);
            this.copySoilData.TabIndex = 24;
            this.copySoilData.Text = "Copy Soil Data to Sheet";
            this.copySoilData.UseVisualStyleBackColor = true;
            this.copySoilData.Click += new System.EventHandler(this.copySoilData_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(6, 134);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(262, 20);
            this.label1.TabIndex = 29;
            this.label1.Text = "Set Data Location on Reference Sheet:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // dispBhRlCell
            // 
            this.dispBhRlCell.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispBhRlCell.Location = new System.Drawing.Point(134, 222);
            this.dispBhRlCell.Name = "dispBhRlCell";
            this.dispBhRlCell.Size = new System.Drawing.Size(134, 20);
            this.dispBhRlCell.TabIndex = 28;
            this.dispBhRlCell.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispBhRlCell.WordWrap = false;
            // 
            // setBhRlCell
            // 
            this.setBhRlCell.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setBhRlCell.Location = new System.Drawing.Point(6, 219);
            this.setBhRlCell.Name = "setBhRlCell";
            this.setBhRlCell.Size = new System.Drawing.Size(122, 25);
            this.setBhRlCell.TabIndex = 27;
            this.setBhRlCell.Text = "Set BH RL Cell";
            this.setBhRlCell.UseVisualStyleBackColor = true;
            // 
            // dispRockStart
            // 
            this.dispRockStart.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRockStart.Location = new System.Drawing.Point(134, 315);
            this.dispRockStart.Name = "dispRockStart";
            this.dispRockStart.Size = new System.Drawing.Size(134, 20);
            this.dispRockStart.TabIndex = 26;
            this.dispRockStart.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRockStart.WordWrap = false;
            // 
            // setRockStart
            // 
            this.setRockStart.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRockStart.Location = new System.Drawing.Point(6, 312);
            this.setRockStart.Name = "setRockStart";
            this.setRockStart.Size = new System.Drawing.Size(122, 25);
            this.setRockStart.TabIndex = 25;
            this.setRockStart.Text = "Set Rock Start Cell";
            this.setRockStart.UseVisualStyleBackColor = true;
            // 
            // dispEffRange
            // 
            this.dispEffRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispEffRange.Location = new System.Drawing.Point(134, 377);
            this.dispEffRange.Name = "dispEffRange";
            this.dispEffRange.Size = new System.Drawing.Size(134, 20);
            this.dispEffRange.TabIndex = 24;
            this.dispEffRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispEffRange.WordWrap = false;
            // 
            // setEffRange
            // 
            this.setEffRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setEffRange.Location = new System.Drawing.Point(6, 374);
            this.setEffRange.Name = "setEffRange";
            this.setEffRange.Size = new System.Drawing.Size(122, 25);
            this.setEffRange.TabIndex = 23;
            this.setEffRange.Text = "Set Eff Range";
            this.setEffRange.UseVisualStyleBackColor = true;
            // 
            // dispQbRange
            // 
            this.dispQbRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispQbRange.Location = new System.Drawing.Point(134, 284);
            this.dispQbRange.Name = "dispQbRange";
            this.dispQbRange.Size = new System.Drawing.Size(134, 20);
            this.dispQbRange.TabIndex = 22;
            this.dispQbRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispQbRange.WordWrap = false;
            // 
            // setQbRange
            // 
            this.setQbRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setQbRange.Location = new System.Drawing.Point(6, 281);
            this.setQbRange.Name = "setQbRange";
            this.setQbRange.Size = new System.Drawing.Size(122, 25);
            this.setQbRange.TabIndex = 21;
            this.setQbRange.Text = "Set qb Range";
            this.setQbRange.UseVisualStyleBackColor = true;
            // 
            // dispFsRange
            // 
            this.dispFsRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFsRange.Location = new System.Drawing.Point(134, 253);
            this.dispFsRange.Name = "dispFsRange";
            this.dispFsRange.Size = new System.Drawing.Size(134, 20);
            this.dispFsRange.TabIndex = 20;
            this.dispFsRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispFsRange.WordWrap = false;
            // 
            // setFsRange
            // 
            this.setFsRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setFsRange.Location = new System.Drawing.Point(6, 250);
            this.setFsRange.Name = "setFsRange";
            this.setFsRange.Size = new System.Drawing.Size(122, 25);
            this.setFsRange.TabIndex = 19;
            this.setFsRange.Text = "Set fs Range";
            this.setFsRange.UseVisualStyleBackColor = true;
            // 
            // dispSoilDest
            // 
            this.dispSoilDest.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispSoilDest.Location = new System.Drawing.Point(134, 191);
            this.dispSoilDest.Name = "dispSoilDest";
            this.dispSoilDest.Size = new System.Drawing.Size(134, 20);
            this.dispSoilDest.TabIndex = 18;
            this.dispSoilDest.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispSoilDest.WordWrap = false;
            // 
            // setSoilDest
            // 
            this.setSoilDest.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSoilDest.Location = new System.Drawing.Point(6, 188);
            this.setSoilDest.Name = "setSoilDest";
            this.setSoilDest.Size = new System.Drawing.Size(122, 25);
            this.setSoilDest.TabIndex = 17;
            this.setSoilDest.Text = "Set Soil Dest.";
            this.setSoilDest.UseVisualStyleBackColor = true;
            // 
            // dispRefSheet
            // 
            this.dispRefSheet.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRefSheet.Location = new System.Drawing.Point(134, 160);
            this.dispRefSheet.Name = "dispRefSheet";
            this.dispRefSheet.Size = new System.Drawing.Size(134, 20);
            this.dispRefSheet.TabIndex = 16;
            this.dispRefSheet.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRefSheet.WordWrap = false;
            // 
            // setRefSheet
            // 
            this.setRefSheet.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRefSheet.Location = new System.Drawing.Point(6, 157);
            this.setRefSheet.Name = "setRefSheet";
            this.setRefSheet.Size = new System.Drawing.Size(122, 25);
            this.setRefSheet.TabIndex = 15;
            this.setRefSheet.Text = "Set Reference Sheet";
            this.setRefSheet.UseVisualStyleBackColor = true;
            // 
            // dispSoilInputData
            // 
            this.dispSoilInputData.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispSoilInputData.Location = new System.Drawing.Point(134, 47);
            this.dispSoilInputData.Name = "dispSoilInputData";
            this.dispSoilInputData.Size = new System.Drawing.Size(134, 20);
            this.dispSoilInputData.TabIndex = 14;
            this.dispSoilInputData.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispSoilInputData.WordWrap = false;
            // 
            // setSoilInputData
            // 
            this.setSoilInputData.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSoilInputData.Location = new System.Drawing.Point(6, 44);
            this.setSoilInputData.Name = "setSoilInputData";
            this.setSoilInputData.Size = new System.Drawing.Size(122, 25);
            this.setSoilInputData.TabIndex = 13;
            this.setSoilInputData.Text = "Set Soil Input Data";
            this.setSoilInputData.UseVisualStyleBackColor = true;
            // 
            // dispSpt100Start
            // 
            this.dispSpt100Start.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispSpt100Start.Location = new System.Drawing.Point(134, 346);
            this.dispSpt100Start.Name = "dispSpt100Start";
            this.dispSpt100Start.Size = new System.Drawing.Size(134, 20);
            this.dispSpt100Start.TabIndex = 40;
            this.dispSpt100Start.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispSpt100Start.WordWrap = false;
            // 
            // setSpt100Start
            // 
            this.setSpt100Start.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSpt100Start.Location = new System.Drawing.Point(6, 343);
            this.setSpt100Start.Name = "setSpt100Start";
            this.setSpt100Start.Size = new System.Drawing.Size(122, 25);
            this.setSpt100Start.TabIndex = 39;
            this.setSpt100Start.Text = "Set SPT100 Start Cell";
            this.setSpt100Start.UseVisualStyleBackColor = true;
            // 
            // PilingPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pilingTabControl);
            this.Name = "PilingPane";
            this.Size = new System.Drawing.Size(300, 1063);
            this.pilingTabControl.ResumeLayout(false);
            this.agsTab.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.agsGroup.ResumeLayout(false);
            this.agsGroup.PerformLayout();
            this.bhGroup.ResumeLayout(false);
            this.bhGroup.PerformLayout();
            this.pilingDesignTab.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl pilingTabControl;
        private System.Windows.Forms.TabPage agsTab;
        private System.Windows.Forms.GroupBox bhGroup;
        private System.Windows.Forms.TextBox DispDrawRange;
        private System.Windows.Forms.Button SetDrawRange;
        private System.Windows.Forms.Button DrawBH;
        private System.Windows.Forms.TextBox DispRockRange;
        private System.Windows.Forms.Button SetRockRange;
        private System.Windows.Forms.GroupBox agsGroup;
        private System.Windows.Forms.Button importAGS;
        private System.Windows.Forms.CheckBox checkRemoveNoSPT;
        private System.Windows.Forms.Button basicAGS;
        private System.Windows.Forms.CheckBox checkFillSoilType;
        private System.Windows.Forms.CheckBox printDescriptionCheck;
        private System.Windows.Forms.TextBox DispNotRockRange;
        private System.Windows.Forms.Button SetNotRockRange;
        private System.Windows.Forms.CheckBox checkDrawRef;
        private System.Windows.Forms.Button removeCont;
        private System.Windows.Forms.TabPage pilingDesignTab;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox dispRockStart;
        private System.Windows.Forms.Button setRockStart;
        private System.Windows.Forms.TextBox dispEffRange;
        private System.Windows.Forms.Button setEffRange;
        private System.Windows.Forms.TextBox dispQbRange;
        private System.Windows.Forms.Button setQbRange;
        private System.Windows.Forms.TextBox dispFsRange;
        private System.Windows.Forms.Button setFsRange;
        private System.Windows.Forms.TextBox dispSoilDest;
        private System.Windows.Forms.Button setSoilDest;
        private System.Windows.Forms.TextBox dispRefSheet;
        private System.Windows.Forms.Button setRefSheet;
        private System.Windows.Forms.TextBox dispSoilInputData;
        private System.Windows.Forms.Button setSoilInputData;
        private System.Windows.Forms.Button copySoilData;
        private System.Windows.Forms.TextBox dispBhRlCell;
        private System.Windows.Forms.Button setBhRlCell;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox dispRockTypeInput;
        private System.Windows.Forms.Button setRockTypeInput;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button setSheetsToRun;
        private System.Windows.Forms.Button designPiles;
        private System.Windows.Forms.CheckBox checkDefaultSPT;
        private System.Windows.Forms.CheckBox checkDeactivateScreen;
        private System.Windows.Forms.TextBox dispNsfTypeInput;
        private System.Windows.Forms.Button setNsfTypeInput;
        private System.Windows.Forms.TextBox dispEfficiencyLower;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox dispEfficiencyUpper;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox dispSpt100Range;
        private System.Windows.Forms.Button setSpt100Range;
        private System.Windows.Forms.Button delSheets;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox checkCompressOutput;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox dispAppendName;
        private System.Windows.Forms.TextBox dispSpt100Start;
        private System.Windows.Forms.Button setSpt100Start;
    }
}
