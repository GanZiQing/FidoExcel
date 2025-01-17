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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.getLoadPatterns = new System.Windows.Forms.Button();
            this.getStoryData = new System.Windows.Forms.Button();
            this.dispJointSortOrder = new System.Windows.Forms.ComboBox();
            this.getJointCoordinates = new System.Windows.Forms.Button();
            this.dispStorySortOrder = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.replaceLoadCheck = new System.Windows.Forms.CheckBox();
            this.setJointDataRange = new System.Windows.Forms.Button();
            this.dispJointDataRange = new System.Windows.Forms.TextBox();
            this.assignWL = new System.Windows.Forms.Button();
            this.dispWindLoadDir = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.calAWL = new System.Windows.Forms.Button();
            this.setStoryRange = new System.Windows.Forms.Button();
            this.dispStoryRange = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.refreshViewCheck = new System.Windows.Forms.CheckBox();
            this.EtabsTabGroup.SuspendLayout();
            this.windLoadPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // EtabsTabGroup
            // 
            this.EtabsTabGroup.Controls.Add(this.windLoadPage);
            this.EtabsTabGroup.Location = new System.Drawing.Point(6, 6);
            this.EtabsTabGroup.Margin = new System.Windows.Forms.Padding(6);
            this.EtabsTabGroup.Name = "EtabsTabGroup";
            this.EtabsTabGroup.SelectedIndex = 0;
            this.EtabsTabGroup.Size = new System.Drawing.Size(539, 1488);
            this.EtabsTabGroup.TabIndex = 4;
            // 
            // windLoadPage
            // 
            this.windLoadPage.BackColor = System.Drawing.SystemColors.Control;
            this.windLoadPage.Controls.Add(this.refreshViewCheck);
            this.windLoadPage.Controls.Add(this.groupBox1);
            this.windLoadPage.Controls.Add(this.groupBox3);
            this.windLoadPage.Location = new System.Drawing.Point(4, 33);
            this.windLoadPage.Name = "windLoadPage";
            this.windLoadPage.Padding = new System.Windows.Forms.Padding(3);
            this.windLoadPage.Size = new System.Drawing.Size(531, 1451);
            this.windLoadPage.TabIndex = 1;
            this.windLoadPage.Text = "Wind Load";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.getLoadPatterns);
            this.groupBox1.Controls.Add(this.getStoryData);
            this.groupBox1.Controls.Add(this.dispJointSortOrder);
            this.groupBox1.Controls.Add(this.getJointCoordinates);
            this.groupBox1.Controls.Add(this.dispStorySortOrder);
            this.groupBox1.Location = new System.Drawing.Point(15, 323);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(502, 207);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Get ETABS Info";
            // 
            // getLoadPatterns
            // 
            this.getLoadPatterns.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getLoadPatterns.Location = new System.Drawing.Point(9, 147);
            this.getLoadPatterns.Margin = new System.Windows.Forms.Padding(6);
            this.getLoadPatterns.Name = "getLoadPatterns";
            this.getLoadPatterns.Size = new System.Drawing.Size(229, 46);
            this.getLoadPatterns.TabIndex = 34;
            this.getLoadPatterns.Text = "Get Load Patterns";
            this.getLoadPatterns.UseVisualStyleBackColor = true;
            this.getLoadPatterns.Click += new System.EventHandler(this.getLoadPatterns_Click);
            // 
            // getStoryData
            // 
            this.getStoryData.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getStoryData.Location = new System.Drawing.Point(9, 31);
            this.getStoryData.Margin = new System.Windows.Forms.Padding(6);
            this.getStoryData.Name = "getStoryData";
            this.getStoryData.Size = new System.Drawing.Size(229, 46);
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
            this.dispJointSortOrder.Location = new System.Drawing.Point(247, 97);
            this.dispJointSortOrder.Name = "dispJointSortOrder";
            this.dispJointSortOrder.Size = new System.Drawing.Size(246, 32);
            this.dispJointSortOrder.TabIndex = 32;
            // 
            // getJointCoordinates
            // 
            this.getJointCoordinates.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getJointCoordinates.Location = new System.Drawing.Point(9, 89);
            this.getJointCoordinates.Margin = new System.Windows.Forms.Padding(6);
            this.getJointCoordinates.Name = "getJointCoordinates";
            this.getJointCoordinates.Size = new System.Drawing.Size(229, 46);
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
            this.dispStorySortOrder.Location = new System.Drawing.Point(247, 45);
            this.dispStorySortOrder.Name = "dispStorySortOrder";
            this.dispStorySortOrder.Size = new System.Drawing.Size(246, 32);
            this.dispStorySortOrder.TabIndex = 33;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.replaceLoadCheck);
            this.groupBox3.Controls.Add(this.setJointDataRange);
            this.groupBox3.Controls.Add(this.dispJointDataRange);
            this.groupBox3.Controls.Add(this.assignWL);
            this.groupBox3.Controls.Add(this.dispWindLoadDir);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.calAWL);
            this.groupBox3.Controls.Add(this.setStoryRange);
            this.groupBox3.Controls.Add(this.dispStoryRange);
            this.groupBox3.Location = new System.Drawing.Point(15, 6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(502, 311);
            this.groupBox3.TabIndex = 29;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Asymmetrical WL";
            // 
            // replaceLoadCheck
            // 
            this.replaceLoadCheck.AutoSize = true;
            this.replaceLoadCheck.Location = new System.Drawing.Point(247, 264);
            this.replaceLoadCheck.Name = "replaceLoadCheck";
            this.replaceLoadCheck.Size = new System.Drawing.Size(146, 28);
            this.replaceLoadCheck.TabIndex = 35;
            this.replaceLoadCheck.Text = "Replace Load";
            this.replaceLoadCheck.UseVisualStyleBackColor = true;
            // 
            // setJointDataRange
            // 
            this.setJointDataRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setJointDataRange.Location = new System.Drawing.Point(9, 196);
            this.setJointDataRange.Margin = new System.Windows.Forms.Padding(6);
            this.setJointDataRange.Name = "setJointDataRange";
            this.setJointDataRange.Size = new System.Drawing.Size(229, 46);
            this.setJointDataRange.TabIndex = 38;
            this.setJointDataRange.Text = "Set Joint Data Range";
            this.setJointDataRange.UseVisualStyleBackColor = true;
            // 
            // dispJointDataRange
            // 
            this.dispJointDataRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispJointDataRange.Location = new System.Drawing.Point(247, 204);
            this.dispJointDataRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispJointDataRange.Name = "dispJointDataRange";
            this.dispJointDataRange.Size = new System.Drawing.Size(246, 29);
            this.dispJointDataRange.TabIndex = 39;
            this.dispJointDataRange.WordWrap = false;
            // 
            // assignWL
            // 
            this.assignWL.ForeColor = System.Drawing.SystemColors.WindowText;
            this.assignWL.Location = new System.Drawing.Point(9, 254);
            this.assignWL.Margin = new System.Windows.Forms.Padding(6);
            this.assignWL.Name = "assignWL";
            this.assignWL.Size = new System.Drawing.Size(229, 46);
            this.assignWL.TabIndex = 37;
            this.assignWL.Text = "Assign WL";
            this.assignWL.UseVisualStyleBackColor = true;
            this.assignWL.Click += new System.EventHandler(this.assignWL_Click);
            // 
            // dispWindLoadDir
            // 
            this.dispWindLoadDir.FormattingEnabled = true;
            this.dispWindLoadDir.Items.AddRange(new object[] {
            "X",
            "Y"});
            this.dispWindLoadDir.Location = new System.Drawing.Point(247, 97);
            this.dispWindLoadDir.Name = "dispWindLoadDir";
            this.dispWindLoadDir.Size = new System.Drawing.Size(246, 32);
            this.dispWindLoadDir.TabIndex = 36;
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(48, 94);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
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
            this.calAWL.Margin = new System.Windows.Forms.Padding(6);
            this.calAWL.Name = "calAWL";
            this.calAWL.Size = new System.Drawing.Size(229, 46);
            this.calAWL.TabIndex = 31;
            this.calAWL.Text = "Calculate AWL";
            this.calAWL.UseVisualStyleBackColor = true;
            this.calAWL.Click += new System.EventHandler(this.calAWL_Click);
            // 
            // setStoryRange
            // 
            this.setStoryRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setStoryRange.Location = new System.Drawing.Point(9, 31);
            this.setStoryRange.Margin = new System.Windows.Forms.Padding(6);
            this.setStoryRange.Name = "setStoryRange";
            this.setStoryRange.Size = new System.Drawing.Size(229, 46);
            this.setStoryRange.TabIndex = 27;
            this.setStoryRange.Text = "Set Story Range";
            this.setStoryRange.UseVisualStyleBackColor = true;
            // 
            // dispStoryRange
            // 
            this.dispStoryRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStoryRange.Location = new System.Drawing.Point(247, 39);
            this.dispStoryRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispStoryRange.Name = "dispStoryRange";
            this.dispStoryRange.Size = new System.Drawing.Size(246, 29);
            this.dispStoryRange.TabIndex = 28;
            this.dispStoryRange.WordWrap = false;
            // 
            // refreshView
            // 
            this.refreshViewCheck.AutoSize = true;
            this.refreshViewCheck.Location = new System.Drawing.Point(24, 536);
            this.refreshViewCheck.Name = "refreshView";
            this.refreshViewCheck.Size = new System.Drawing.Size(141, 28);
            this.refreshViewCheck.TabIndex = 40;
            this.refreshViewCheck.Text = "Refresh View";
            this.refreshViewCheck.UseVisualStyleBackColor = true;
            // 
            // ETABSTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.EtabsTabGroup);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "ETABSTaskPane";
            this.Size = new System.Drawing.Size(550, 1500);
            this.EtabsTabGroup.ResumeLayout(false);
            this.windLoadPage.ResumeLayout(false);
            this.windLoadPage.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
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
        private System.Windows.Forms.Button assignWL;
        private System.Windows.Forms.Button getLoadPatterns;
        private System.Windows.Forms.Button setJointDataRange;
        private System.Windows.Forms.TextBox dispJointDataRange;
        private System.Windows.Forms.CheckBox replaceLoadCheck;
        private System.Windows.Forms.CheckBox refreshViewCheck;
    }
}
