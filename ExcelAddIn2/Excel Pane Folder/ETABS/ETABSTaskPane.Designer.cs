namespace ExcelAddIn2
{
    partial class ETABSTaskPane
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
            this.EtabsTabGroup = new System.Windows.Forms.TabControl();
            this.windLoadPage = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.replaceLoadCheck = new System.Windows.Forms.CheckBox();
            this.refreshViewCheck = new System.Windows.Forms.CheckBox();
            this.setJointDataRange = new System.Windows.Forms.Button();
            this.dispJointDataRange = new System.Windows.Forms.TextBox();
            this.assignWL = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.getLoadPatterns = new System.Windows.Forms.Button();
            this.getStoryData = new System.Windows.Forms.Button();
            this.dispJointSortOrder = new System.Windows.Forms.ComboBox();
            this.getJointCoordinates = new System.Windows.Forms.Button();
            this.dispStorySortOrder = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dispWindLoadDir = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.calAWL = new System.Windows.Forms.Button();
            this.setStoryRange = new System.Windows.Forms.Button();
            this.dispStoryRange = new System.Windows.Forms.TextBox();
            this.baseShearPage = new System.Windows.Forms.TabPage();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.dispTableFormat = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.checkObjectsAreUnique = new System.Windows.Forms.Button();
            this.getAllReactionsButt = new System.Windows.Forms.Button();
            this.testGetObjButt = new System.Windows.Forms.Button();
            this.getActiveLoadComboButt = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.getBaseReactionButt = new System.Windows.Forms.Button();
            this.dispGetForcesObj = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.setLcRange = new System.Windows.Forms.Button();
            this.dispLcRange = new System.Windows.Forms.TextBox();
            this.setGroupRange = new System.Windows.Forms.Button();
            this.dispGroupRange = new System.Windows.Forms.TextBox();
            this.utilitiesPage = new System.Windows.Forms.TabPage();
            this.selectGroupBox = new System.Windows.Forms.GroupBox();
            this.getPilingForces = new System.Windows.Forms.Button();
            this.getWallUNBut = new System.Windows.Forms.Button();
            this.getWallPierBut = new System.Windows.Forms.Button();
            this.setWallPierBut = new System.Windows.Forms.Button();
            this.errorGroupBox = new System.Windows.Forms.GroupBox();
            this.groupAndImportLog = new System.Windows.Forms.Button();
            this.findSlantedWalls = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.openWRN = new System.Windows.Forms.Button();
            this.groupAndImportWrn = new System.Windows.Forms.Button();
            this.openLog = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.replicateBySpacingDisp = new System.Windows.Forms.Button();
            this.replicateByDispBut = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.getGroupNames = new System.Windows.Forms.Button();
            this.EtabsTabGroup.SuspendLayout();
            this.windLoadPage.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.baseShearPage.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.utilitiesPage.SuspendLayout();
            this.selectGroupBox.SuspendLayout();
            this.errorGroupBox.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // EtabsTabGroup
            // 
            this.EtabsTabGroup.Controls.Add(this.windLoadPage);
            this.EtabsTabGroup.Controls.Add(this.baseShearPage);
            this.EtabsTabGroup.Controls.Add(this.utilitiesPage);
            this.EtabsTabGroup.Location = new System.Drawing.Point(5, 6);
            this.EtabsTabGroup.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.EtabsTabGroup.Name = "EtabsTabGroup";
            this.EtabsTabGroup.SelectedIndex = 0;
            this.EtabsTabGroup.Size = new System.Drawing.Size(539, 1488);
            this.EtabsTabGroup.TabIndex = 4;
            // 
            // windLoadPage
            // 
            this.windLoadPage.BackColor = System.Drawing.SystemColors.Control;
            this.windLoadPage.Controls.Add(this.groupBox2);
            this.windLoadPage.Controls.Add(this.groupBox1);
            this.windLoadPage.Controls.Add(this.groupBox3);
            this.windLoadPage.Location = new System.Drawing.Point(4, 33);
            this.windLoadPage.Margin = new System.Windows.Forms.Padding(4);
            this.windLoadPage.Name = "windLoadPage";
            this.windLoadPage.Padding = new System.Windows.Forms.Padding(4);
            this.windLoadPage.Size = new System.Drawing.Size(531, 1451);
            this.windLoadPage.TabIndex = 1;
            this.windLoadPage.Text = "Wind Load";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.replaceLoadCheck);
            this.groupBox2.Controls.Add(this.refreshViewCheck);
            this.groupBox2.Controls.Add(this.setJointDataRange);
            this.groupBox2.Controls.Add(this.dispJointDataRange);
            this.groupBox2.Controls.Add(this.assignWL);
            this.groupBox2.Location = new System.Drawing.Point(10, 430);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.groupBox2.Size = new System.Drawing.Size(501, 186);
            this.groupBox2.TabIndex = 41;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Assign Wind Load";
            // 
            // replaceLoadCheck
            // 
            this.replaceLoadCheck.AutoSize = true;
            this.replaceLoadCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.replaceLoadCheck.Location = new System.Drawing.Point(244, 109);
            this.replaceLoadCheck.Margin = new System.Windows.Forms.Padding(4);
            this.replaceLoadCheck.Name = "replaceLoadCheck";
            this.replaceLoadCheck.Size = new System.Drawing.Size(146, 28);
            this.replaceLoadCheck.TabIndex = 40;
            this.replaceLoadCheck.Text = "Replace Load";
            this.replaceLoadCheck.UseVisualStyleBackColor = true;
            // 
            // refreshViewCheck
            // 
            this.refreshViewCheck.AutoSize = true;
            this.refreshViewCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.refreshViewCheck.Location = new System.Drawing.Point(244, 148);
            this.refreshViewCheck.Margin = new System.Windows.Forms.Padding(4);
            this.refreshViewCheck.Name = "refreshViewCheck";
            this.refreshViewCheck.Size = new System.Drawing.Size(141, 28);
            this.refreshViewCheck.TabIndex = 40;
            this.refreshViewCheck.Text = "Refresh View";
            this.refreshViewCheck.UseVisualStyleBackColor = true;
            // 
            // setJointDataRange
            // 
            this.setJointDataRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setJointDataRange.Location = new System.Drawing.Point(5, 41);
            this.setJointDataRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.setJointDataRange.Name = "setJointDataRange";
            this.setJointDataRange.Size = new System.Drawing.Size(230, 46);
            this.setJointDataRange.TabIndex = 42;
            this.setJointDataRange.Text = "Set Joint Data Range";
            this.setJointDataRange.UseVisualStyleBackColor = true;
            // 
            // dispJointDataRange
            // 
            this.dispJointDataRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispJointDataRange.Location = new System.Drawing.Point(244, 50);
            this.dispJointDataRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.dispJointDataRange.Name = "dispJointDataRange";
            this.dispJointDataRange.Size = new System.Drawing.Size(246, 29);
            this.dispJointDataRange.TabIndex = 43;
            this.dispJointDataRange.WordWrap = false;
            // 
            // assignWL
            // 
            this.assignWL.ForeColor = System.Drawing.SystemColors.WindowText;
            this.assignWL.Location = new System.Drawing.Point(5, 100);
            this.assignWL.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.assignWL.Name = "assignWL";
            this.assignWL.Size = new System.Drawing.Size(230, 46);
            this.assignWL.TabIndex = 41;
            this.assignWL.Text = "Assign WL";
            this.assignWL.UseVisualStyleBackColor = true;
            this.assignWL.Click += new System.EventHandler(this.assignWL_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.getLoadPatterns);
            this.groupBox1.Controls.Add(this.getStoryData);
            this.groupBox1.Controls.Add(this.dispJointSortOrder);
            this.groupBox1.Controls.Add(this.getJointCoordinates);
            this.groupBox1.Controls.Add(this.dispStorySortOrder);
            this.groupBox1.Location = new System.Drawing.Point(7, 7);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(502, 206);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Get ETABS Info";
            // 
            // getLoadPatterns
            // 
            this.getLoadPatterns.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getLoadPatterns.Location = new System.Drawing.Point(10, 148);
            this.getLoadPatterns.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getLoadPatterns.Name = "getLoadPatterns";
            this.getLoadPatterns.Size = new System.Drawing.Size(230, 46);
            this.getLoadPatterns.TabIndex = 34;
            this.getLoadPatterns.Text = "Get Load Patterns";
            this.getLoadPatterns.UseVisualStyleBackColor = true;
            this.getLoadPatterns.Click += new System.EventHandler(this.getLoadPatterns_Click);
            // 
            // getStoryData
            // 
            this.getStoryData.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getStoryData.Location = new System.Drawing.Point(10, 31);
            this.getStoryData.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getStoryData.Name = "getStoryData";
            this.getStoryData.Size = new System.Drawing.Size(230, 46);
            this.getStoryData.TabIndex = 30;
            this.getStoryData.Text = "Get Story Data";
            this.getStoryData.UseVisualStyleBackColor = true;
            this.getStoryData.Click += new System.EventHandler(this.getStoryData_Click);
            // 
            // dispJointSortOrder
            // 
            this.dispJointSortOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.dispJointSortOrder.FormattingEnabled = true;
            this.dispJointSortOrder.Items.AddRange(new object[] {
            "Z, X, Y",
            "Z, Y, X",
            "X, Y, Z",
            "X, Z, Y",
            "Y, X, Z",
            "Y, Z, X"});
            this.dispJointSortOrder.Location = new System.Drawing.Point(247, 98);
            this.dispJointSortOrder.Margin = new System.Windows.Forms.Padding(4);
            this.dispJointSortOrder.Name = "dispJointSortOrder";
            this.dispJointSortOrder.Size = new System.Drawing.Size(246, 32);
            this.dispJointSortOrder.TabIndex = 32;
            // 
            // getJointCoordinates
            // 
            this.getJointCoordinates.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getJointCoordinates.Location = new System.Drawing.Point(10, 89);
            this.getJointCoordinates.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getJointCoordinates.Name = "getJointCoordinates";
            this.getJointCoordinates.Size = new System.Drawing.Size(230, 46);
            this.getJointCoordinates.TabIndex = 29;
            this.getJointCoordinates.Text = "Get Joint Coordinates";
            this.getJointCoordinates.UseVisualStyleBackColor = true;
            this.getJointCoordinates.Click += new System.EventHandler(this.getJointCoordinates_Click);
            // 
            // dispStorySortOrder
            // 
            this.dispStorySortOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.dispStorySortOrder.FormattingEnabled = true;
            this.dispStorySortOrder.Items.AddRange(new object[] {
            "Top to Bottom",
            "Bottom to Top"});
            this.dispStorySortOrder.Location = new System.Drawing.Point(247, 44);
            this.dispStorySortOrder.Margin = new System.Windows.Forms.Padding(4);
            this.dispStorySortOrder.Name = "dispStorySortOrder";
            this.dispStorySortOrder.Size = new System.Drawing.Size(246, 32);
            this.dispStorySortOrder.TabIndex = 33;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dispWindLoadDir);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.calAWL);
            this.groupBox3.Controls.Add(this.setStoryRange);
            this.groupBox3.Controls.Add(this.dispStoryRange);
            this.groupBox3.Location = new System.Drawing.Point(7, 222);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox3.Size = new System.Drawing.Size(502, 199);
            this.groupBox3.TabIndex = 29;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Calculate WL";
            // 
            // dispWindLoadDir
            // 
            this.dispWindLoadDir.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.dispWindLoadDir.FormattingEnabled = true;
            this.dispWindLoadDir.Items.AddRange(new object[] {
            "X",
            "Y"});
            this.dispWindLoadDir.Location = new System.Drawing.Point(247, 98);
            this.dispWindLoadDir.Margin = new System.Windows.Forms.Padding(4);
            this.dispWindLoadDir.Name = "dispWindLoadDir";
            this.dispWindLoadDir.Size = new System.Drawing.Size(246, 32);
            this.dispWindLoadDir.TabIndex = 36;
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(48, 94);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(147, 37);
            this.label1.TabIndex = 35;
            this.label1.Text = "Wind Load Dir. ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // calAWL
            // 
            this.calAWL.ForeColor = System.Drawing.SystemColors.WindowText;
            this.calAWL.Location = new System.Drawing.Point(136, 138);
            this.calAWL.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.calAWL.Name = "calAWL";
            this.calAWL.Size = new System.Drawing.Size(230, 46);
            this.calAWL.TabIndex = 31;
            this.calAWL.Text = "Calculate AWL";
            this.calAWL.UseVisualStyleBackColor = true;
            this.calAWL.Click += new System.EventHandler(this.calAWL_Click);
            // 
            // setStoryRange
            // 
            this.setStoryRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setStoryRange.Location = new System.Drawing.Point(10, 31);
            this.setStoryRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.setStoryRange.Name = "setStoryRange";
            this.setStoryRange.Size = new System.Drawing.Size(230, 46);
            this.setStoryRange.TabIndex = 27;
            this.setStoryRange.Text = "Set Story Range";
            this.setStoryRange.UseVisualStyleBackColor = true;
            // 
            // dispStoryRange
            // 
            this.dispStoryRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStoryRange.Location = new System.Drawing.Point(247, 38);
            this.dispStoryRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.dispStoryRange.Name = "dispStoryRange";
            this.dispStoryRange.Size = new System.Drawing.Size(246, 29);
            this.dispStoryRange.TabIndex = 28;
            this.dispStoryRange.WordWrap = false;
            // 
            // baseShearPage
            // 
            this.baseShearPage.BackColor = System.Drawing.SystemColors.Control;
            this.baseShearPage.Controls.Add(this.groupBox6);
            this.baseShearPage.Controls.Add(this.label3);
            this.baseShearPage.Controls.Add(this.groupBox7);
            this.baseShearPage.Controls.Add(this.groupBox5);
            this.baseShearPage.Location = new System.Drawing.Point(4, 33);
            this.baseShearPage.Margin = new System.Windows.Forms.Padding(4);
            this.baseShearPage.Name = "baseShearPage";
            this.baseShearPage.Padding = new System.Windows.Forms.Padding(4);
            this.baseShearPage.Size = new System.Drawing.Size(531, 1451);
            this.baseShearPage.TabIndex = 3;
            this.baseShearPage.Text = "Base Shear";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.dispTableFormat);
            this.groupBox6.Controls.Add(this.label4);
            this.groupBox6.Location = new System.Drawing.Point(7, 549);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(502, 82);
            this.groupBox6.TabIndex = 43;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Settings";
            // 
            // dispTableFormat
            // 
            this.dispTableFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.dispTableFormat.FormattingEnabled = true;
            this.dispTableFormat.Items.AddRange(new object[] {
            "Append Right",
            "Append Bottom"});
            this.dispTableFormat.Location = new System.Drawing.Point(246, 33);
            this.dispTableFormat.Margin = new System.Windows.Forms.Padding(4);
            this.dispTableFormat.Name = "dispTableFormat";
            this.dispTableFormat.Size = new System.Drawing.Size(246, 32);
            this.dispTableFormat.TabIndex = 40;
            // 
            // label4
            // 
            this.label4.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label4.Location = new System.Drawing.Point(9, 25);
            this.label4.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(230, 46);
            this.label4.TabIndex = 39;
            this.label4.Text = "Table Format";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label3.Location = new System.Drawing.Point(9, 471);
            this.label3.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(500, 75);
            this.label3.TabIndex = 39;
            this.label3.Text = "Note: \r\nEnsure that desired load cases/combos are selected in ETABS Table \"Choose" +
    " Tables For Display\"";
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.getGroupNames);
            this.groupBox7.Controls.Add(this.checkObjectsAreUnique);
            this.groupBox7.Controls.Add(this.getAllReactionsButt);
            this.groupBox7.Controls.Add(this.testGetObjButt);
            this.groupBox7.Controls.Add(this.getActiveLoadComboButt);
            this.groupBox7.Location = new System.Drawing.Point(7, 264);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(502, 204);
            this.groupBox7.TabIndex = 38;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Checks";
            // 
            // checkObjectsAreUnique
            // 
            this.checkObjectsAreUnique.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkObjectsAreUnique.Location = new System.Drawing.Point(22, 147);
            this.checkObjectsAreUnique.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.checkObjectsAreUnique.Name = "checkObjectsAreUnique";
            this.checkObjectsAreUnique.Size = new System.Drawing.Size(230, 46);
            this.checkObjectsAreUnique.TabIndex = 42;
            this.checkObjectsAreUnique.Text = "Check Col/Wall Unique";
            this.checkObjectsAreUnique.UseVisualStyleBackColor = true;
            this.checkObjectsAreUnique.Click += new System.EventHandler(this.checkObjectsAreUnique_Click);
            // 
            // getAllReactionsButt
            // 
            this.getAllReactionsButt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getAllReactionsButt.Location = new System.Drawing.Point(22, 89);
            this.getAllReactionsButt.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getAllReactionsButt.Name = "getAllReactionsButt";
            this.getAllReactionsButt.Size = new System.Drawing.Size(230, 46);
            this.getAllReactionsButt.TabIndex = 41;
            this.getAllReactionsButt.Text = "Get All Reactions";
            this.getAllReactionsButt.UseVisualStyleBackColor = true;
            this.getAllReactionsButt.Click += new System.EventHandler(this.getAllReactionsButt_Click);
            // 
            // testGetObjButt
            // 
            this.testGetObjButt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.testGetObjButt.Location = new System.Drawing.Point(262, 89);
            this.testGetObjButt.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.testGetObjButt.Name = "testGetObjButt";
            this.testGetObjButt.Size = new System.Drawing.Size(230, 46);
            this.testGetObjButt.TabIndex = 39;
            this.testGetObjButt.Text = "Get Joints Considered";
            this.testGetObjButt.UseVisualStyleBackColor = true;
            this.testGetObjButt.Click += new System.EventHandler(this.getJointsUsedButt_Click);
            // 
            // getActiveLoadComboButt
            // 
            this.getActiveLoadComboButt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getActiveLoadComboButt.Location = new System.Drawing.Point(262, 31);
            this.getActiveLoadComboButt.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getActiveLoadComboButt.Name = "getActiveLoadComboButt";
            this.getActiveLoadComboButt.Size = new System.Drawing.Size(230, 46);
            this.getActiveLoadComboButt.TabIndex = 40;
            this.getActiveLoadComboButt.Text = "Get Active Load Combos";
            this.getActiveLoadComboButt.UseVisualStyleBackColor = true;
            this.getActiveLoadComboButt.Click += new System.EventHandler(this.getActiveLoadComboButt_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.getBaseReactionButt);
            this.groupBox5.Controls.Add(this.dispGetForcesObj);
            this.groupBox5.Controls.Add(this.label2);
            this.groupBox5.Controls.Add(this.setLcRange);
            this.groupBox5.Controls.Add(this.dispLcRange);
            this.groupBox5.Controls.Add(this.setGroupRange);
            this.groupBox5.Controls.Add(this.dispGroupRange);
            this.groupBox5.Location = new System.Drawing.Point(7, 7);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox5.Size = new System.Drawing.Size(502, 250);
            this.groupBox5.TabIndex = 36;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Settings";
            // 
            // getBaseReactionButt
            // 
            this.getBaseReactionButt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getBaseReactionButt.Location = new System.Drawing.Point(147, 192);
            this.getBaseReactionButt.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getBaseReactionButt.Name = "getBaseReactionButt";
            this.getBaseReactionButt.Size = new System.Drawing.Size(230, 46);
            this.getBaseReactionButt.TabIndex = 29;
            this.getBaseReactionButt.Text = "Get Base Reaction";
            this.getBaseReactionButt.UseVisualStyleBackColor = true;
            this.getBaseReactionButt.Click += new System.EventHandler(this.getBaseShearButt_Click);
            // 
            // dispGetForcesObj
            // 
            this.dispGetForcesObj.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.dispGetForcesObj.FormattingEnabled = true;
            this.dispGetForcesObj.Items.AddRange(new object[] {
            "Column & Wall",
            "Joints"});
            this.dispGetForcesObj.Location = new System.Drawing.Point(246, 150);
            this.dispGetForcesObj.Margin = new System.Windows.Forms.Padding(4);
            this.dispGetForcesObj.Name = "dispGetForcesObj";
            this.dispGetForcesObj.Size = new System.Drawing.Size(246, 32);
            this.dispGetForcesObj.TabIndex = 38;
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label2.Location = new System.Drawing.Point(9, 142);
            this.label2.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(230, 46);
            this.label2.TabIndex = 37;
            this.label2.Text = "Get Forces From";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // setLcRange
            // 
            this.setLcRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setLcRange.Location = new System.Drawing.Point(9, 90);
            this.setLcRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.setLcRange.Name = "setLcRange";
            this.setLcRange.Size = new System.Drawing.Size(230, 46);
            this.setLcRange.TabIndex = 31;
            this.setLcRange.Text = "Set Load Combo Range";
            this.setLcRange.UseVisualStyleBackColor = true;
            // 
            // dispLcRange
            // 
            this.dispLcRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispLcRange.Location = new System.Drawing.Point(246, 98);
            this.dispLcRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.dispLcRange.Name = "dispLcRange";
            this.dispLcRange.Size = new System.Drawing.Size(246, 29);
            this.dispLcRange.TabIndex = 32;
            this.dispLcRange.WordWrap = false;
            // 
            // setGroupRange
            // 
            this.setGroupRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setGroupRange.Location = new System.Drawing.Point(9, 32);
            this.setGroupRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.setGroupRange.Name = "setGroupRange";
            this.setGroupRange.Size = new System.Drawing.Size(230, 46);
            this.setGroupRange.TabIndex = 29;
            this.setGroupRange.Text = "Set Group Range";
            this.setGroupRange.UseVisualStyleBackColor = true;
            // 
            // dispGroupRange
            // 
            this.dispGroupRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispGroupRange.Location = new System.Drawing.Point(247, 40);
            this.dispGroupRange.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.dispGroupRange.Name = "dispGroupRange";
            this.dispGroupRange.Size = new System.Drawing.Size(246, 29);
            this.dispGroupRange.TabIndex = 30;
            this.dispGroupRange.WordWrap = false;
            // 
            // utilitiesPage
            // 
            this.utilitiesPage.BackColor = System.Drawing.SystemColors.Control;
            this.utilitiesPage.Controls.Add(this.selectGroupBox);
            this.utilitiesPage.Controls.Add(this.errorGroupBox);
            this.utilitiesPage.Controls.Add(this.groupBox4);
            this.utilitiesPage.Location = new System.Drawing.Point(4, 33);
            this.utilitiesPage.Margin = new System.Windows.Forms.Padding(4);
            this.utilitiesPage.Name = "utilitiesPage";
            this.utilitiesPage.Padding = new System.Windows.Forms.Padding(4);
            this.utilitiesPage.Size = new System.Drawing.Size(531, 1451);
            this.utilitiesPage.TabIndex = 2;
            this.utilitiesPage.Text = "Utilities";
            // 
            // selectGroupBox
            // 
            this.selectGroupBox.Controls.Add(this.getPilingForces);
            this.selectGroupBox.Controls.Add(this.getWallUNBut);
            this.selectGroupBox.Controls.Add(this.getWallPierBut);
            this.selectGroupBox.Controls.Add(this.setWallPierBut);
            this.selectGroupBox.Location = new System.Drawing.Point(7, 498);
            this.selectGroupBox.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.selectGroupBox.Name = "selectGroupBox";
            this.selectGroupBox.Padding = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.selectGroupBox.Size = new System.Drawing.Size(501, 275);
            this.selectGroupBox.TabIndex = 47;
            this.selectGroupBox.TabStop = false;
            this.selectGroupBox.Text = "Get, Select, Sets";
            // 
            // getPilingForces
            // 
            this.getPilingForces.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getPilingForces.Location = new System.Drawing.Point(246, 217);
            this.getPilingForces.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getPilingForces.Name = "getPilingForces";
            this.getPilingForces.Size = new System.Drawing.Size(230, 46);
            this.getPilingForces.TabIndex = 46;
            this.getPilingForces.Text = "Get Piling Forces";
            this.getPilingForces.UseVisualStyleBackColor = true;
            this.getPilingForces.Click += new System.EventHandler(this.getPilingForces_Click);
            // 
            // getWallUNBut
            // 
            this.getWallUNBut.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getWallUNBut.Location = new System.Drawing.Point(11, 35);
            this.getWallUNBut.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getWallUNBut.Name = "getWallUNBut";
            this.getWallUNBut.Size = new System.Drawing.Size(230, 46);
            this.getWallUNBut.TabIndex = 43;
            this.getWallUNBut.Text = "Get Wall UN";
            this.getWallUNBut.UseVisualStyleBackColor = true;
            this.getWallUNBut.Click += new System.EventHandler(this.getWallUNBut_Click);
            // 
            // getWallPierBut
            // 
            this.getWallPierBut.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getWallPierBut.Location = new System.Drawing.Point(11, 94);
            this.getWallPierBut.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getWallPierBut.Name = "getWallPierBut";
            this.getWallPierBut.Size = new System.Drawing.Size(230, 46);
            this.getWallPierBut.TabIndex = 44;
            this.getWallPierBut.Text = "Get Wall Pier";
            this.getWallPierBut.UseVisualStyleBackColor = true;
            this.getWallPierBut.Click += new System.EventHandler(this.getWallPierBut_Click);
            // 
            // setWallPierBut
            // 
            this.setWallPierBut.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setWallPierBut.Location = new System.Drawing.Point(11, 151);
            this.setWallPierBut.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.setWallPierBut.Name = "setWallPierBut";
            this.setWallPierBut.Size = new System.Drawing.Size(230, 46);
            this.setWallPierBut.TabIndex = 45;
            this.setWallPierBut.Text = "Set Wall Pier";
            this.setWallPierBut.UseVisualStyleBackColor = true;
            this.setWallPierBut.Click += new System.EventHandler(this.setWallPierBut_Click);
            // 
            // errorGroupBox
            // 
            this.errorGroupBox.Controls.Add(this.groupAndImportLog);
            this.errorGroupBox.Controls.Add(this.findSlantedWalls);
            this.errorGroupBox.Controls.Add(this.textBox2);
            this.errorGroupBox.Controls.Add(this.openWRN);
            this.errorGroupBox.Controls.Add(this.groupAndImportWrn);
            this.errorGroupBox.Controls.Add(this.openLog);
            this.errorGroupBox.Controls.Add(this.textBox1);
            this.errorGroupBox.Controls.Add(this.textBox3);
            this.errorGroupBox.Location = new System.Drawing.Point(7, 170);
            this.errorGroupBox.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.errorGroupBox.Name = "errorGroupBox";
            this.errorGroupBox.Padding = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.errorGroupBox.Size = new System.Drawing.Size(501, 318);
            this.errorGroupBox.TabIndex = 46;
            this.errorGroupBox.TabStop = false;
            this.errorGroupBox.Text = "Error Checking";
            // 
            // groupAndImportLog
            // 
            this.groupAndImportLog.ForeColor = System.Drawing.SystemColors.WindowText;
            this.groupAndImportLog.Location = new System.Drawing.Point(11, 162);
            this.groupAndImportLog.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.groupAndImportLog.Name = "groupAndImportLog";
            this.groupAndImportLog.Size = new System.Drawing.Size(230, 46);
            this.groupAndImportLog.TabIndex = 43;
            this.groupAndImportLog.Text = "Import .LOG";
            this.groupAndImportLog.UseVisualStyleBackColor = true;
            this.groupAndImportLog.Click += new System.EventHandler(this.groupAndImportLog_Click);
            // 
            // findSlantedWalls
            // 
            this.findSlantedWalls.ForeColor = System.Drawing.SystemColors.WindowText;
            this.findSlantedWalls.Location = new System.Drawing.Point(11, 254);
            this.findSlantedWalls.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.findSlantedWalls.Name = "findSlantedWalls";
            this.findSlantedWalls.Size = new System.Drawing.Size(230, 46);
            this.findSlantedWalls.TabIndex = 44;
            this.findSlantedWalls.Text = "Find Slanted Walls";
            this.findSlantedWalls.UseVisualStyleBackColor = true;
            this.findSlantedWalls.Click += new System.EventHandler(this.findSlantedWalls_Click);
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Control;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Location = new System.Drawing.Point(11, 220);
            this.textBox2.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(183, 22);
            this.textBox2.TabIndex = 51;
            this.textBox2.TabStop = false;
            this.textBox2.Text = "Others";
            // 
            // openWRN
            // 
            this.openWRN.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openWRN.Location = new System.Drawing.Point(260, 70);
            this.openWRN.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.openWRN.Name = "openWRN";
            this.openWRN.Size = new System.Drawing.Size(230, 46);
            this.openWRN.TabIndex = 50;
            this.openWRN.Text = "Open Wrn";
            this.openWRN.UseVisualStyleBackColor = true;
            this.openWRN.Click += new System.EventHandler(this.openWRN_Click);
            // 
            // groupAndImportWrn
            // 
            this.groupAndImportWrn.ForeColor = System.Drawing.SystemColors.WindowText;
            this.groupAndImportWrn.Location = new System.Drawing.Point(11, 70);
            this.groupAndImportWrn.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.groupAndImportWrn.Name = "groupAndImportWrn";
            this.groupAndImportWrn.Size = new System.Drawing.Size(230, 46);
            this.groupAndImportWrn.TabIndex = 49;
            this.groupAndImportWrn.Text = "Import .WRN";
            this.groupAndImportWrn.UseVisualStyleBackColor = true;
            this.groupAndImportWrn.Click += new System.EventHandler(this.groupAndImportWrn_Click);
            // 
            // openLog
            // 
            this.openLog.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openLog.Location = new System.Drawing.Point(260, 162);
            this.openLog.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.openLog.Name = "openLog";
            this.openLog.Size = new System.Drawing.Size(230, 46);
            this.openLog.TabIndex = 48;
            this.openLog.Text = "Open Log";
            this.openLog.UseVisualStyleBackColor = true;
            this.openLog.Click += new System.EventHandler(this.openLog_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(11, 127);
            this.textBox1.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(220, 22);
            this.textBox1.TabIndex = 47;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "Analysis Errors (.LOG)";
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.SystemColors.Control;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Location = new System.Drawing.Point(11, 35);
            this.textBox3.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(220, 22);
            this.textBox3.TabIndex = 46;
            this.textBox3.TabStop = false;
            this.textBox3.Text = "Check Model (.WRN)";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.replicateBySpacingDisp);
            this.groupBox4.Controls.Add(this.replicateByDispBut);
            this.groupBox4.Location = new System.Drawing.Point(7, 7);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox4.Size = new System.Drawing.Size(502, 154);
            this.groupBox4.TabIndex = 35;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Replicate";
            // 
            // replicateBySpacingDisp
            // 
            this.replicateBySpacingDisp.ForeColor = System.Drawing.SystemColors.WindowText;
            this.replicateBySpacingDisp.Location = new System.Drawing.Point(10, 89);
            this.replicateBySpacingDisp.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.replicateBySpacingDisp.Name = "replicateBySpacingDisp";
            this.replicateBySpacingDisp.Size = new System.Drawing.Size(230, 46);
            this.replicateBySpacingDisp.TabIndex = 42;
            this.replicateBySpacingDisp.Text = "Replicate by Spacing";
            this.replicateBySpacingDisp.UseVisualStyleBackColor = true;
            this.replicateBySpacingDisp.Click += new System.EventHandler(this.replicateBySpacingDisp_Click);
            // 
            // replicateByDispBut
            // 
            this.replicateByDispBut.ForeColor = System.Drawing.SystemColors.WindowText;
            this.replicateByDispBut.Location = new System.Drawing.Point(10, 31);
            this.replicateByDispBut.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.replicateByDispBut.Name = "replicateByDispBut";
            this.replicateByDispBut.Size = new System.Drawing.Size(230, 46);
            this.replicateByDispBut.TabIndex = 30;
            this.replicateByDispBut.Text = "Replicate by Disp.";
            this.replicateByDispBut.UseVisualStyleBackColor = true;
            this.replicateByDispBut.Click += new System.EventHandler(this.replicateByDispBut_Click);
            // 
            // getGroupNames
            // 
            this.getGroupNames.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getGroupNames.Location = new System.Drawing.Point(22, 31);
            this.getGroupNames.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.getGroupNames.Name = "getGroupNames";
            this.getGroupNames.Size = new System.Drawing.Size(230, 46);
            this.getGroupNames.TabIndex = 43;
            this.getGroupNames.Text = "Get Group Names";
            this.getGroupNames.UseVisualStyleBackColor = true;
            this.getGroupNames.Click += new System.EventHandler(this.getGroupNames_Click);
            // 
            // ETABSTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.EtabsTabGroup);
            this.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.Name = "ETABSTaskPane";
            this.Size = new System.Drawing.Size(550, 1499);
            this.EtabsTabGroup.ResumeLayout(false);
            this.windLoadPage.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.baseShearPage.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.utilitiesPage.ResumeLayout(false);
            this.selectGroupBox.ResumeLayout(false);
            this.errorGroupBox.ResumeLayout(false);
            this.errorGroupBox.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl EtabsTabGroup;
        private System.Windows.Forms.TabPage windLoadPage;
        private System.Windows.Forms.Button setStoryRange;
        private System.Windows.Forms.TextBox dispStoryRange;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button getJointCoordinates;
        private System.Windows.Forms.Button getStoryData;
        private System.Windows.Forms.Button calAWL;
        private System.Windows.Forms.ComboBox dispJointSortOrder;
        private System.Windows.Forms.ComboBox dispStorySortOrder;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox dispWindLoadDir;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button getLoadPatterns;
        private System.Windows.Forms.CheckBox refreshViewCheck;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox replaceLoadCheck;
        private System.Windows.Forms.Button setJointDataRange;
        private System.Windows.Forms.TextBox dispJointDataRange;
        private System.Windows.Forms.Button assignWL;
        private System.Windows.Forms.TabPage utilitiesPage;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button replicateByDispBut;
        private System.Windows.Forms.Button replicateBySpacingDisp;
        private System.Windows.Forms.Button setWallPierBut;
        private System.Windows.Forms.Button getWallPierBut;
        private System.Windows.Forms.Button getWallUNBut;
        private System.Windows.Forms.Button findSlantedWalls;
        private System.Windows.Forms.GroupBox errorGroupBox;
        private System.Windows.Forms.Button groupAndImportLog;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Button openLog;
        private System.Windows.Forms.GroupBox selectGroupBox;
        private System.Windows.Forms.Button openWRN;
        private System.Windows.Forms.Button groupAndImportWrn;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button getPilingForces;
        private System.Windows.Forms.TabPage baseShearPage;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button setLcRange;
        private System.Windows.Forms.TextBox dispLcRange;
        private System.Windows.Forms.Button setGroupRange;
        private System.Windows.Forms.TextBox dispGroupRange;
        private System.Windows.Forms.Button getBaseReactionButt;
        private System.Windows.Forms.ComboBox dispGetForcesObj;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button testGetObjButt;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.Button getActiveLoadComboButt;
        private System.Windows.Forms.Button getAllReactionsButt;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button checkObjectsAreUnique;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.ComboBox dispTableFormat;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button getGroupNames;
    }
}
