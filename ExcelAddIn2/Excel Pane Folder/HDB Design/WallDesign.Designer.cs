namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class WallDesign
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
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.wallDesignPage = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.unifyChangesButt = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.resetFontColourCheckSheetCheck = new System.Windows.Forms.CheckBox();
            this.resetFontColourRebarTableCheck = new System.Windows.Forms.CheckBox();
            this.backupSheetCheck = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dispMaxAs = new System.Windows.Forms.TextBox();
            this.designRebar = new System.Windows.Forms.Button();
            this.overwriteRebarCheck = new System.Windows.Forms.CheckBox();
            this.unifyChangesCheck = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dispTargetUR = new System.Windows.Forms.TextBox();
            this.setRebarHeirarchy = new System.Windows.Forms.Button();
            this.dispRebarHeirarchy = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.setStatusCol = new System.Windows.Forms.Button();
            this.dispStatusCol = new System.Windows.Forms.TextBox();
            this.matchWallRebar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.setMatchStoreyCol = new System.Windows.Forms.Button();
            this.dispMatchStoreyCol = new System.Windows.Forms.TextBox();
            this.setOutputCol = new System.Windows.Forms.Button();
            this.dispOutputCol = new System.Windows.Forms.TextBox();
            this.setPierLabelRange = new System.Windows.Forms.Button();
            this.dispPierLabelRange = new System.Windows.Forms.TextBox();
            this.setStoreyTable = new System.Windows.Forms.Button();
            this.dispStoreyTable = new System.Windows.Forms.TextBox();
            this.setRebarTable = new System.Windows.Forms.Button();
            this.dispRebarTable = new System.Windows.Forms.TextBox();
            this.unmergerTabPage = new System.Windows.Forms.TabPage();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.decomposeTable = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.setDecomposeRange = new System.Windows.Forms.Button();
            this.dispDecomposeRange = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tabControl1.SuspendLayout();
            this.wallDesignPage.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.unmergerTabPage.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.wallDesignPage);
            this.tabControl1.Controls.Add(this.unmergerTabPage);
            this.tabControl1.Location = new System.Drawing.Point(4, 4);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(546, 1525);
            this.tabControl1.TabIndex = 1;
            // 
            // wallDesignPage
            // 
            this.wallDesignPage.BackColor = System.Drawing.SystemColors.Control;
            this.wallDesignPage.Controls.Add(this.groupBox4);
            this.wallDesignPage.Controls.Add(this.groupBox3);
            this.wallDesignPage.Controls.Add(this.groupBox2);
            this.wallDesignPage.Controls.Add(this.groupBox1);
            this.wallDesignPage.Location = new System.Drawing.Point(4, 33);
            this.wallDesignPage.Margin = new System.Windows.Forms.Padding(4);
            this.wallDesignPage.Name = "wallDesignPage";
            this.wallDesignPage.Padding = new System.Windows.Forms.Padding(4);
            this.wallDesignPage.Size = new System.Drawing.Size(538, 1488);
            this.wallDesignPage.TabIndex = 0;
            this.wallDesignPage.Text = "Wall Design";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.unifyChangesButt);
            this.groupBox4.Location = new System.Drawing.Point(9, 844);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox4.Size = new System.Drawing.Size(521, 100);
            this.groupBox4.TabIndex = 42;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Unify Changes";
            // 
            // unifyChangesButt
            // 
            this.unifyChangesButt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.unifyChangesButt.Location = new System.Drawing.Point(138, 31);
            this.unifyChangesButt.Margin = new System.Windows.Forms.Padding(6);
            this.unifyChangesButt.Name = "unifyChangesButt";
            this.unifyChangesButt.Size = new System.Drawing.Size(229, 46);
            this.unifyChangesButt.TabIndex = 54;
            this.unifyChangesButt.Text = "Unify Changes";
            this.unifyChangesButt.UseVisualStyleBackColor = true;
            this.unifyChangesButt.Click += new System.EventHandler(this.unifyChangesButt_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.resetFontColourCheckSheetCheck);
            this.groupBox3.Controls.Add(this.resetFontColourRebarTableCheck);
            this.groupBox3.Controls.Add(this.backupSheetCheck);
            this.groupBox3.Location = new System.Drawing.Point(9, 951);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox3.Size = new System.Drawing.Size(521, 146);
            this.groupBox3.TabIndex = 41;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Additional Settings";
            // 
            // resetFontColourCheckSheetCheck
            // 
            this.resetFontColourCheckSheetCheck.AutoSize = true;
            this.resetFontColourCheckSheetCheck.Location = new System.Drawing.Point(24, 102);
            this.resetFontColourCheckSheetCheck.Margin = new System.Windows.Forms.Padding(4);
            this.resetFontColourCheckSheetCheck.Name = "resetFontColourCheckSheetCheck";
            this.resetFontColourCheckSheetCheck.Size = new System.Drawing.Size(306, 28);
            this.resetFontColourCheckSheetCheck.TabIndex = 58;
            this.resetFontColourCheckSheetCheck.Text = "Reset Font Colour (Check Sheet)";
            this.resetFontColourCheckSheetCheck.UseVisualStyleBackColor = true;
            // 
            // resetFontColourRebarTableCheck
            // 
            this.resetFontColourRebarTableCheck.AutoSize = true;
            this.resetFontColourRebarTableCheck.Location = new System.Drawing.Point(24, 66);
            this.resetFontColourRebarTableCheck.Margin = new System.Windows.Forms.Padding(4);
            this.resetFontColourRebarTableCheck.Name = "resetFontColourRebarTableCheck";
            this.resetFontColourRebarTableCheck.Size = new System.Drawing.Size(302, 28);
            this.resetFontColourRebarTableCheck.TabIndex = 57;
            this.resetFontColourRebarTableCheck.Text = "Reset Font Colour (Rebar Table)";
            this.resetFontColourRebarTableCheck.UseVisualStyleBackColor = true;
            // 
            // backupSheetCheck
            // 
            this.backupSheetCheck.AutoSize = true;
            this.backupSheetCheck.Location = new System.Drawing.Point(24, 28);
            this.backupSheetCheck.Margin = new System.Windows.Forms.Padding(4);
            this.backupSheetCheck.Name = "backupSheetCheck";
            this.backupSheetCheck.Size = new System.Drawing.Size(206, 28);
            this.backupSheetCheck.TabIndex = 56;
            this.backupSheetCheck.Text = "Create Backup Sheet";
            this.backupSheetCheck.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.dispMaxAs);
            this.groupBox2.Controls.Add(this.designRebar);
            this.groupBox2.Controls.Add(this.overwriteRebarCheck);
            this.groupBox2.Controls.Add(this.unifyChangesCheck);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.dispTargetUR);
            this.groupBox2.Controls.Add(this.setRebarHeirarchy);
            this.groupBox2.Controls.Add(this.dispRebarHeirarchy);
            this.groupBox2.Location = new System.Drawing.Point(9, 524);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(521, 312);
            this.groupBox2.TabIndex = 40;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Modify Rebars";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label3.Location = new System.Drawing.Point(18, 133);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(235, 37);
            this.label3.TabIndex = 56;
            this.label3.Text = "Max %As:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dispMaxAs
            // 
            this.dispMaxAs.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispMaxAs.Location = new System.Drawing.Point(273, 135);
            this.dispMaxAs.Margin = new System.Windows.Forms.Padding(6);
            this.dispMaxAs.Name = "dispMaxAs";
            this.dispMaxAs.Size = new System.Drawing.Size(224, 29);
            this.dispMaxAs.TabIndex = 55;
            this.dispMaxAs.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispMaxAs.WordWrap = false;
            // 
            // designRebar
            // 
            this.designRebar.ForeColor = System.Drawing.SystemColors.WindowText;
            this.designRebar.Location = new System.Drawing.Point(139, 253);
            this.designRebar.Margin = new System.Windows.Forms.Padding(6);
            this.designRebar.Name = "designRebar";
            this.designRebar.Size = new System.Drawing.Size(229, 46);
            this.designRebar.TabIndex = 39;
            this.designRebar.Text = "Design Rebar";
            this.designRebar.UseVisualStyleBackColor = true;
            this.designRebar.Click += new System.EventHandler(this.designRebar_Click);
            // 
            // overwriteRebarCheck
            // 
            this.overwriteRebarCheck.AutoSize = true;
            this.overwriteRebarCheck.Location = new System.Drawing.Point(24, 174);
            this.overwriteRebarCheck.Margin = new System.Windows.Forms.Padding(4);
            this.overwriteRebarCheck.Name = "overwriteRebarCheck";
            this.overwriteRebarCheck.Size = new System.Drawing.Size(211, 28);
            this.overwriteRebarCheck.TabIndex = 54;
            this.overwriteRebarCheck.Text = "Overwrite Initial Rebar";
            this.overwriteRebarCheck.UseVisualStyleBackColor = true;
            // 
            // unifyChangesCheck
            // 
            this.unifyChangesCheck.AutoSize = true;
            this.unifyChangesCheck.Location = new System.Drawing.Point(24, 212);
            this.unifyChangesCheck.Margin = new System.Windows.Forms.Padding(4);
            this.unifyChangesCheck.Name = "unifyChangesCheck";
            this.unifyChangesCheck.Size = new System.Drawing.Size(258, 28);
            this.unifyChangesCheck.TabIndex = 55;
            this.unifyChangesCheck.Text = "Unify Changes After Design";
            this.unifyChangesCheck.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label2.Location = new System.Drawing.Point(18, 83);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(235, 37);
            this.label2.TabIndex = 51;
            this.label2.Text = "Target UR:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dispTargetUR
            // 
            this.dispTargetUR.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispTargetUR.Location = new System.Drawing.Point(273, 87);
            this.dispTargetUR.Margin = new System.Windows.Forms.Padding(6);
            this.dispTargetUR.Name = "dispTargetUR";
            this.dispTargetUR.Size = new System.Drawing.Size(224, 29);
            this.dispTargetUR.TabIndex = 50;
            this.dispTargetUR.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispTargetUR.WordWrap = false;
            // 
            // setRebarHeirarchy
            // 
            this.setRebarHeirarchy.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRebarHeirarchy.Location = new System.Drawing.Point(24, 31);
            this.setRebarHeirarchy.Margin = new System.Windows.Forms.Padding(6);
            this.setRebarHeirarchy.Name = "setRebarHeirarchy";
            this.setRebarHeirarchy.Size = new System.Drawing.Size(229, 46);
            this.setRebarHeirarchy.TabIndex = 45;
            this.setRebarHeirarchy.Text = "Set Rebar Heirarchy";
            this.setRebarHeirarchy.UseVisualStyleBackColor = true;
            // 
            // dispRebarHeirarchy
            // 
            this.dispRebarHeirarchy.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRebarHeirarchy.Location = new System.Drawing.Point(273, 37);
            this.dispRebarHeirarchy.Margin = new System.Windows.Forms.Padding(6);
            this.dispRebarHeirarchy.Name = "dispRebarHeirarchy";
            this.dispRebarHeirarchy.Size = new System.Drawing.Size(224, 29);
            this.dispRebarHeirarchy.TabIndex = 46;
            this.dispRebarHeirarchy.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRebarHeirarchy.WordWrap = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.setStatusCol);
            this.groupBox1.Controls.Add(this.dispStatusCol);
            this.groupBox1.Controls.Add(this.matchWallRebar);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.setMatchStoreyCol);
            this.groupBox1.Controls.Add(this.dispMatchStoreyCol);
            this.groupBox1.Controls.Add(this.setOutputCol);
            this.groupBox1.Controls.Add(this.dispOutputCol);
            this.groupBox1.Controls.Add(this.setPierLabelRange);
            this.groupBox1.Controls.Add(this.dispPierLabelRange);
            this.groupBox1.Controls.Add(this.setStoreyTable);
            this.groupBox1.Controls.Add(this.dispStoreyTable);
            this.groupBox1.Controls.Add(this.setRebarTable);
            this.groupBox1.Controls.Add(this.dispRebarTable);
            this.groupBox1.Location = new System.Drawing.Point(9, 6);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(521, 511);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Match Rebar";
            // 
            // setStatusCol
            // 
            this.setStatusCol.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setStatusCol.Location = new System.Drawing.Point(24, 391);
            this.setStatusCol.Margin = new System.Windows.Forms.Padding(6);
            this.setStatusCol.Name = "setStatusCol";
            this.setStatusCol.Size = new System.Drawing.Size(229, 46);
            this.setStatusCol.TabIndex = 43;
            this.setStatusCol.Text = "Set Status Col";
            this.setStatusCol.UseVisualStyleBackColor = true;
            // 
            // dispStatusCol
            // 
            this.dispStatusCol.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStatusCol.Location = new System.Drawing.Point(273, 397);
            this.dispStatusCol.Margin = new System.Windows.Forms.Padding(6);
            this.dispStatusCol.Name = "dispStatusCol";
            this.dispStatusCol.Size = new System.Drawing.Size(224, 29);
            this.dispStatusCol.TabIndex = 44;
            this.dispStatusCol.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispStatusCol.WordWrap = false;
            // 
            // matchWallRebar
            // 
            this.matchWallRebar.ForeColor = System.Drawing.SystemColors.WindowText;
            this.matchWallRebar.Location = new System.Drawing.Point(139, 449);
            this.matchWallRebar.Margin = new System.Windows.Forms.Padding(6);
            this.matchWallRebar.Name = "matchWallRebar";
            this.matchWallRebar.Size = new System.Drawing.Size(229, 46);
            this.matchWallRebar.TabIndex = 38;
            this.matchWallRebar.Text = "Match Rebar";
            this.matchWallRebar.UseVisualStyleBackColor = true;
            this.matchWallRebar.Click += new System.EventHandler(this.matchWallRebar_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(9, 179);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(480, 37);
            this.label1.TabIndex = 42;
            this.label1.Text = "Match Table:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label5.Location = new System.Drawing.Point(9, 26);
            this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(480, 37);
            this.label5.TabIndex = 36;
            this.label5.Text = "Reference Tables:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // setMatchStoreyCol
            // 
            this.setMatchStoreyCol.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setMatchStoreyCol.Location = new System.Drawing.Point(24, 275);
            this.setMatchStoreyCol.Margin = new System.Windows.Forms.Padding(6);
            this.setMatchStoreyCol.Name = "setMatchStoreyCol";
            this.setMatchStoreyCol.Size = new System.Drawing.Size(229, 46);
            this.setMatchStoreyCol.TabIndex = 40;
            this.setMatchStoreyCol.Text = "Set Storey Col";
            this.setMatchStoreyCol.UseVisualStyleBackColor = true;
            // 
            // dispMatchStoreyCol
            // 
            this.dispMatchStoreyCol.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispMatchStoreyCol.Location = new System.Drawing.Point(273, 281);
            this.dispMatchStoreyCol.Margin = new System.Windows.Forms.Padding(6);
            this.dispMatchStoreyCol.Name = "dispMatchStoreyCol";
            this.dispMatchStoreyCol.Size = new System.Drawing.Size(224, 29);
            this.dispMatchStoreyCol.TabIndex = 41;
            this.dispMatchStoreyCol.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispMatchStoreyCol.WordWrap = false;
            // 
            // setOutputCol
            // 
            this.setOutputCol.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setOutputCol.Location = new System.Drawing.Point(24, 332);
            this.setOutputCol.Margin = new System.Windows.Forms.Padding(6);
            this.setOutputCol.Name = "setOutputCol";
            this.setOutputCol.Size = new System.Drawing.Size(229, 46);
            this.setOutputCol.TabIndex = 38;
            this.setOutputCol.Text = "Set Output Col";
            this.setOutputCol.UseVisualStyleBackColor = true;
            // 
            // dispOutputCol
            // 
            this.dispOutputCol.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOutputCol.Location = new System.Drawing.Point(273, 338);
            this.dispOutputCol.Margin = new System.Windows.Forms.Padding(6);
            this.dispOutputCol.Name = "dispOutputCol";
            this.dispOutputCol.Size = new System.Drawing.Size(224, 29);
            this.dispOutputCol.TabIndex = 39;
            this.dispOutputCol.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispOutputCol.WordWrap = false;
            // 
            // setPierLabelRange
            // 
            this.setPierLabelRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPierLabelRange.Location = new System.Drawing.Point(24, 218);
            this.setPierLabelRange.Margin = new System.Windows.Forms.Padding(6);
            this.setPierLabelRange.Name = "setPierLabelRange";
            this.setPierLabelRange.Size = new System.Drawing.Size(229, 46);
            this.setPierLabelRange.TabIndex = 36;
            this.setPierLabelRange.Text = "Set Pier Label Range";
            this.setPierLabelRange.UseVisualStyleBackColor = true;
            // 
            // dispPierLabelRange
            // 
            this.dispPierLabelRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPierLabelRange.Location = new System.Drawing.Point(273, 222);
            this.dispPierLabelRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispPierLabelRange.Name = "dispPierLabelRange";
            this.dispPierLabelRange.Size = new System.Drawing.Size(224, 29);
            this.dispPierLabelRange.TabIndex = 37;
            this.dispPierLabelRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispPierLabelRange.WordWrap = false;
            // 
            // setStoreyTable
            // 
            this.setStoreyTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setStoreyTable.Location = new System.Drawing.Point(24, 127);
            this.setStoreyTable.Margin = new System.Windows.Forms.Padding(6);
            this.setStoreyTable.Name = "setStoreyTable";
            this.setStoreyTable.Size = new System.Drawing.Size(229, 46);
            this.setStoreyTable.TabIndex = 34;
            this.setStoreyTable.Text = "Set Storey Table";
            this.setStoreyTable.UseVisualStyleBackColor = true;
            // 
            // dispStoreyTable
            // 
            this.dispStoreyTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStoreyTable.Location = new System.Drawing.Point(273, 133);
            this.dispStoreyTable.Margin = new System.Windows.Forms.Padding(6);
            this.dispStoreyTable.Name = "dispStoreyTable";
            this.dispStoreyTable.Size = new System.Drawing.Size(224, 29);
            this.dispStoreyTable.TabIndex = 35;
            this.dispStoreyTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispStoreyTable.WordWrap = false;
            // 
            // setRebarTable
            // 
            this.setRebarTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRebarTable.Location = new System.Drawing.Point(24, 68);
            this.setRebarTable.Margin = new System.Windows.Forms.Padding(6);
            this.setRebarTable.Name = "setRebarTable";
            this.setRebarTable.Size = new System.Drawing.Size(229, 46);
            this.setRebarTable.TabIndex = 32;
            this.setRebarTable.Text = "Set Rebar Table";
            this.setRebarTable.UseVisualStyleBackColor = true;
            // 
            // dispRebarTable
            // 
            this.dispRebarTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRebarTable.Location = new System.Drawing.Point(273, 76);
            this.dispRebarTable.Margin = new System.Windows.Forms.Padding(6);
            this.dispRebarTable.Name = "dispRebarTable";
            this.dispRebarTable.Size = new System.Drawing.Size(224, 29);
            this.dispRebarTable.TabIndex = 33;
            this.dispRebarTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRebarTable.WordWrap = false;
            // 
            // unmergerTabPage
            // 
            this.unmergerTabPage.Controls.Add(this.groupBox5);
            this.unmergerTabPage.Location = new System.Drawing.Point(4, 33);
            this.unmergerTabPage.Name = "unmergerTabPage";
            this.unmergerTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.unmergerTabPage.Size = new System.Drawing.Size(538, 1488);
            this.unmergerTabPage.TabIndex = 1;
            this.unmergerTabPage.Text = "Unmerger";
            this.unmergerTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.decomposeTable);
            this.groupBox5.Controls.Add(this.label6);
            this.groupBox5.Controls.Add(this.setDecomposeRange);
            this.groupBox5.Controls.Add(this.dispDecomposeRange);
            this.groupBox5.Location = new System.Drawing.Point(7, 7);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox5.Size = new System.Drawing.Size(521, 184);
            this.groupBox5.TabIndex = 35;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Match Rebar";
            // 
            // decomposeTable
            // 
            this.decomposeTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.decomposeTable.Location = new System.Drawing.Point(129, 126);
            this.decomposeTable.Margin = new System.Windows.Forms.Padding(6);
            this.decomposeTable.Name = "decomposeTable";
            this.decomposeTable.Size = new System.Drawing.Size(229, 46);
            this.decomposeTable.TabIndex = 38;
            this.decomposeTable.Text = "Decompose Table";
            this.decomposeTable.UseVisualStyleBackColor = true;
            this.decomposeTable.Click += new System.EventHandler(this.decomposeTable_Click);
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label6.Location = new System.Drawing.Point(9, 26);
            this.label6.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(480, 37);
            this.label6.TabIndex = 36;
            this.label6.Text = "Reference Tables:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // setDecomposeRange
            // 
            this.setDecomposeRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setDecomposeRange.Location = new System.Drawing.Point(24, 68);
            this.setDecomposeRange.Margin = new System.Windows.Forms.Padding(6);
            this.setDecomposeRange.Name = "setDecomposeRange";
            this.setDecomposeRange.Size = new System.Drawing.Size(229, 46);
            this.setDecomposeRange.TabIndex = 32;
            this.setDecomposeRange.Text = "Set Table";
            this.setDecomposeRange.UseVisualStyleBackColor = true;
            // 
            // dispDecomposeRange
            // 
            this.dispDecomposeRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispDecomposeRange.Location = new System.Drawing.Point(273, 76);
            this.dispDecomposeRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispDecomposeRange.Name = "dispDecomposeRange";
            this.dispDecomposeRange.Size = new System.Drawing.Size(224, 29);
            this.dispDecomposeRange.TabIndex = 33;
            this.dispDecomposeRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispDecomposeRange.WordWrap = false;
            // 
            // WallDesign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "WallDesign";
            this.Size = new System.Drawing.Size(550, 1532);
            this.tabControl1.ResumeLayout(false);
            this.wallDesignPage.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.unmergerTabPage.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage wallDesignPage;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button matchWallRebar;
        private System.Windows.Forms.Button setPierLabelRange;
        private System.Windows.Forms.TextBox dispPierLabelRange;
        private System.Windows.Forms.Button setStoreyTable;
        private System.Windows.Forms.TextBox dispStoreyTable;
        private System.Windows.Forms.Button setRebarTable;
        private System.Windows.Forms.TextBox dispRebarTable;
        private System.Windows.Forms.Button setOutputCol;
        private System.Windows.Forms.TextBox dispOutputCol;
        private System.Windows.Forms.Button setMatchStoreyCol;
        private System.Windows.Forms.TextBox dispMatchStoreyCol;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button setStatusCol;
        private System.Windows.Forms.TextBox dispStatusCol;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button designRebar;
        private System.Windows.Forms.Button setRebarHeirarchy;
        private System.Windows.Forms.TextBox dispRebarHeirarchy;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox dispTargetUR;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button unifyChangesButt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox dispMaxAs;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox resetFontColourRebarTableCheck;
        private System.Windows.Forms.CheckBox backupSheetCheck;
        private System.Windows.Forms.CheckBox unifyChangesCheck;
        private System.Windows.Forms.CheckBox overwriteRebarCheck;
        private System.Windows.Forms.CheckBox resetFontColourCheckSheetCheck;
        private System.Windows.Forms.TabPage unmergerTabPage;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button decomposeTable;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button setDecomposeRange;
        private System.Windows.Forms.TextBox dispDecomposeRange;
    }
}
