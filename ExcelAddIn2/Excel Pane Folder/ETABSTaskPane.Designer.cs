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
            this.refreshViewCheck = new System.Windows.Forms.CheckBox();
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
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.replaceLoadCheck = new System.Windows.Forms.CheckBox();
            this.setJointDataRange = new System.Windows.Forms.Button();
            this.dispJointDataRange = new System.Windows.Forms.TextBox();
            this.assignWL = new System.Windows.Forms.Button();
            this.EtabsTabGroup.SuspendLayout();
            this.windLoadPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // EtabsTabGroup
            // 
            this.EtabsTabGroup.Controls.Add(this.windLoadPage);
            this.EtabsTabGroup.Location = new System.Drawing.Point(3, 3);
            this.EtabsTabGroup.Name = "EtabsTabGroup";
            this.EtabsTabGroup.SelectedIndex = 0;
            this.EtabsTabGroup.Size = new System.Drawing.Size(294, 806);
            this.EtabsTabGroup.TabIndex = 4;
            // 
            // windLoadPage
            // 
            this.windLoadPage.BackColor = System.Drawing.SystemColors.Control;
            this.windLoadPage.Controls.Add(this.groupBox2);
            this.windLoadPage.Controls.Add(this.groupBox1);
            this.windLoadPage.Controls.Add(this.groupBox3);
            this.windLoadPage.Location = new System.Drawing.Point(4, 22);
            this.windLoadPage.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.windLoadPage.Name = "windLoadPage";
            this.windLoadPage.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.windLoadPage.Size = new System.Drawing.Size(286, 780);
            this.windLoadPage.TabIndex = 1;
            this.windLoadPage.Text = "Wind Load";
            // 
            // refreshViewCheck
            // 
            this.refreshViewCheck.AutoSize = true;
            this.refreshViewCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.refreshViewCheck.Location = new System.Drawing.Point(133, 80);
            this.refreshViewCheck.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.refreshViewCheck.Name = "refreshViewCheck";
            this.refreshViewCheck.Size = new System.Drawing.Size(89, 17);
            this.refreshViewCheck.TabIndex = 40;
            this.refreshViewCheck.Text = "Refresh View";
            this.refreshViewCheck.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.getLoadPatterns);
            this.groupBox1.Controls.Add(this.getStoryData);
            this.groupBox1.Controls.Add(this.dispJointSortOrder);
            this.groupBox1.Controls.Add(this.getJointCoordinates);
            this.groupBox1.Controls.Add(this.dispStorySortOrder);
            this.groupBox1.Location = new System.Drawing.Point(4, 4);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(274, 112);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Get ETABS Info";
            // 
            // getLoadPatterns
            // 
            this.getLoadPatterns.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getLoadPatterns.Location = new System.Drawing.Point(5, 80);
            this.getLoadPatterns.Name = "getLoadPatterns";
            this.getLoadPatterns.Size = new System.Drawing.Size(125, 25);
            this.getLoadPatterns.TabIndex = 34;
            this.getLoadPatterns.Text = "Get Load Patterns";
            this.getLoadPatterns.UseVisualStyleBackColor = true;
            this.getLoadPatterns.Click += new System.EventHandler(this.getLoadPatterns_Click);
            // 
            // getStoryData
            // 
            this.getStoryData.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getStoryData.Location = new System.Drawing.Point(5, 17);
            this.getStoryData.Name = "getStoryData";
            this.getStoryData.Size = new System.Drawing.Size(125, 25);
            this.getStoryData.TabIndex = 30;
            this.getStoryData.Text = "Get Story Data";
            this.getStoryData.UseVisualStyleBackColor = true;
            this.getStoryData.Click += new System.EventHandler(this.getStoryData_Click);
            // 
            // dispJointSortOrder
            // 
            this.dispJointSortOrder.FormattingEnabled = true;
            this.dispJointSortOrder.Items.AddRange(new object[] {
            "Z, X, Y",
            "Z, Y, X",
            "X, Y, Z",
            "X, Z, Y",
            "Y, X, Z",
            "Y, Z, X"});
            this.dispJointSortOrder.Location = new System.Drawing.Point(135, 53);
            this.dispJointSortOrder.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dispJointSortOrder.Name = "dispJointSortOrder";
            this.dispJointSortOrder.Size = new System.Drawing.Size(136, 21);
            this.dispJointSortOrder.TabIndex = 32;
            // 
            // getJointCoordinates
            // 
            this.getJointCoordinates.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getJointCoordinates.Location = new System.Drawing.Point(5, 48);
            this.getJointCoordinates.Name = "getJointCoordinates";
            this.getJointCoordinates.Size = new System.Drawing.Size(125, 25);
            this.getJointCoordinates.TabIndex = 29;
            this.getJointCoordinates.Text = "Get Joint Coordinates";
            this.getJointCoordinates.UseVisualStyleBackColor = true;
            this.getJointCoordinates.Click += new System.EventHandler(this.getJointCoordinates_Click);
            // 
            // dispStorySortOrder
            // 
            this.dispStorySortOrder.FormattingEnabled = true;
            this.dispStorySortOrder.Items.AddRange(new object[] {
            "Top to Bottom",
            "Bottom to Top"});
            this.dispStorySortOrder.Location = new System.Drawing.Point(135, 24);
            this.dispStorySortOrder.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dispStorySortOrder.Name = "dispStorySortOrder";
            this.dispStorySortOrder.Size = new System.Drawing.Size(136, 21);
            this.dispStorySortOrder.TabIndex = 33;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dispWindLoadDir);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.calAWL);
            this.groupBox3.Controls.Add(this.setStoryRange);
            this.groupBox3.Controls.Add(this.dispStoryRange);
            this.groupBox3.Location = new System.Drawing.Point(4, 120);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Size = new System.Drawing.Size(274, 108);
            this.groupBox3.TabIndex = 29;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Calculate WL";
            // 
            // dispWindLoadDir
            // 
            this.dispWindLoadDir.FormattingEnabled = true;
            this.dispWindLoadDir.Items.AddRange(new object[] {
            "X",
            "Y"});
            this.dispWindLoadDir.Location = new System.Drawing.Point(135, 53);
            this.dispWindLoadDir.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dispWindLoadDir.Name = "dispWindLoadDir";
            this.dispWindLoadDir.Size = new System.Drawing.Size(136, 21);
            this.dispWindLoadDir.TabIndex = 36;
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(26, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 20);
            this.label1.TabIndex = 35;
            this.label1.Text = "Wind Load Dir. ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // calAWL
            // 
            this.calAWL.ForeColor = System.Drawing.SystemColors.WindowText;
            this.calAWL.Location = new System.Drawing.Point(74, 75);
            this.calAWL.Name = "calAWL";
            this.calAWL.Size = new System.Drawing.Size(125, 25);
            this.calAWL.TabIndex = 31;
            this.calAWL.Text = "Calculate AWL";
            this.calAWL.UseVisualStyleBackColor = true;
            this.calAWL.Click += new System.EventHandler(this.calAWL_Click);
            // 
            // setStoryRange
            // 
            this.setStoryRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setStoryRange.Location = new System.Drawing.Point(5, 17);
            this.setStoryRange.Name = "setStoryRange";
            this.setStoryRange.Size = new System.Drawing.Size(125, 25);
            this.setStoryRange.TabIndex = 27;
            this.setStoryRange.Text = "Set Story Range";
            this.setStoryRange.UseVisualStyleBackColor = true;
            // 
            // dispStoryRange
            // 
            this.dispStoryRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStoryRange.Location = new System.Drawing.Point(135, 21);
            this.dispStoryRange.Name = "dispStoryRange";
            this.dispStoryRange.Size = new System.Drawing.Size(136, 20);
            this.dispStoryRange.TabIndex = 28;
            this.dispStoryRange.WordWrap = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.replaceLoadCheck);
            this.groupBox2.Controls.Add(this.refreshViewCheck);
            this.groupBox2.Controls.Add(this.setJointDataRange);
            this.groupBox2.Controls.Add(this.dispJointDataRange);
            this.groupBox2.Controls.Add(this.assignWL);
            this.groupBox2.Location = new System.Drawing.Point(5, 233);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(273, 101);
            this.groupBox2.TabIndex = 41;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Assign Wind Load";
            // 
            // replaceLoadCheck
            // 
            this.replaceLoadCheck.AutoSize = true;
            this.replaceLoadCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.replaceLoadCheck.Location = new System.Drawing.Point(133, 59);
            this.replaceLoadCheck.Margin = new System.Windows.Forms.Padding(2);
            this.replaceLoadCheck.Name = "replaceLoadCheck";
            this.replaceLoadCheck.Size = new System.Drawing.Size(93, 17);
            this.replaceLoadCheck.TabIndex = 40;
            this.replaceLoadCheck.Text = "Replace Load";
            this.replaceLoadCheck.UseVisualStyleBackColor = true;
            // 
            // setJointDataRange
            // 
            this.setJointDataRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setJointDataRange.Location = new System.Drawing.Point(3, 22);
            this.setJointDataRange.Name = "setJointDataRange";
            this.setJointDataRange.Size = new System.Drawing.Size(125, 25);
            this.setJointDataRange.TabIndex = 42;
            this.setJointDataRange.Text = "Set Joint Data Range";
            this.setJointDataRange.UseVisualStyleBackColor = true;
            // 
            // dispJointDataRange
            // 
            this.dispJointDataRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispJointDataRange.Location = new System.Drawing.Point(133, 27);
            this.dispJointDataRange.Name = "dispJointDataRange";
            this.dispJointDataRange.Size = new System.Drawing.Size(136, 20);
            this.dispJointDataRange.TabIndex = 43;
            this.dispJointDataRange.WordWrap = false;
            // 
            // assignWL
            // 
            this.assignWL.ForeColor = System.Drawing.SystemColors.WindowText;
            this.assignWL.Location = new System.Drawing.Point(3, 54);
            this.assignWL.Name = "assignWL";
            this.assignWL.Size = new System.Drawing.Size(125, 25);
            this.assignWL.TabIndex = 41;
            this.assignWL.Text = "Assign WL";
            this.assignWL.UseVisualStyleBackColor = true;
            // 
            // ETABSTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.EtabsTabGroup);
            this.Name = "ETABSTaskPane";
            this.Size = new System.Drawing.Size(300, 812);
            this.EtabsTabGroup.ResumeLayout(false);
            this.windLoadPage.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
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
    }
}
