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
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.getStoreyNames = new System.Windows.Forms.Button();
            this.getJointCoordinates = new System.Windows.Forms.Button();
            this.setStoreyRange = new System.Windows.Forms.Button();
            this.dispStoreyRange = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.calAWL = new System.Windows.Forms.Button();
            this.EtabsTabGroup.SuspendLayout();
            this.windLoadPage.SuspendLayout();
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
            this.EtabsTabGroup.Size = new System.Drawing.Size(539, 1911);
            this.EtabsTabGroup.TabIndex = 4;
            // 
            // windLoadPage
            // 
            this.windLoadPage.Controls.Add(this.groupBox3);
            this.windLoadPage.Location = new System.Drawing.Point(4, 33);
            this.windLoadPage.Name = "windLoadPage";
            this.windLoadPage.Padding = new System.Windows.Forms.Padding(3);
            this.windLoadPage.Size = new System.Drawing.Size(531, 1874);
            this.windLoadPage.TabIndex = 1;
            this.windLoadPage.Text = "Wind Load";
            this.windLoadPage.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.calAWL);
            this.groupBox3.Controls.Add(this.getStoreyNames);
            this.groupBox3.Controls.Add(this.getJointCoordinates);
            this.groupBox3.Controls.Add(this.setStoreyRange);
            this.groupBox3.Controls.Add(this.dispStoreyRange);
            this.groupBox3.Location = new System.Drawing.Point(6, 6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(519, 347);
            this.groupBox3.TabIndex = 29;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Settings";
            // 
            // getStoreyNames
            // 
            this.getStoreyNames.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getStoreyNames.Location = new System.Drawing.Point(9, 31);
            this.getStoreyNames.Margin = new System.Windows.Forms.Padding(6);
            this.getStoreyNames.Name = "getStoreyNames";
            this.getStoreyNames.Size = new System.Drawing.Size(229, 46);
            this.getStoreyNames.TabIndex = 30;
            this.getStoreyNames.Text = "Get Storey Names";
            this.getStoreyNames.UseVisualStyleBackColor = true;
            this.getStoreyNames.Click += new System.EventHandler(this.getStoreyNames_Click);
            // 
            // getJointCoordinates
            // 
            this.getJointCoordinates.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getJointCoordinates.Location = new System.Drawing.Point(9, 147);
            this.getJointCoordinates.Margin = new System.Windows.Forms.Padding(6);
            this.getJointCoordinates.Name = "getJointCoordinates";
            this.getJointCoordinates.Size = new System.Drawing.Size(229, 46);
            this.getJointCoordinates.TabIndex = 29;
            this.getJointCoordinates.Text = "Get Joint Coordinates";
            this.getJointCoordinates.UseVisualStyleBackColor = true;
            this.getJointCoordinates.Click += new System.EventHandler(this.getJointCoordinates_Click);
            // 
            // setStoreyRange
            // 
            this.setStoreyRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setStoreyRange.Location = new System.Drawing.Point(9, 89);
            this.setStoreyRange.Margin = new System.Windows.Forms.Padding(6);
            this.setStoreyRange.Name = "setStoreyRange";
            this.setStoreyRange.Size = new System.Drawing.Size(229, 46);
            this.setStoreyRange.TabIndex = 27;
            this.setStoreyRange.Text = "Set Storey Range";
            this.setStoreyRange.UseVisualStyleBackColor = true;
            // 
            // dispStoreyRange
            // 
            this.dispStoreyRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStoreyRange.Location = new System.Drawing.Point(260, 95);
            this.dispStoreyRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispStoreyRange.Name = "dispStoreyRange";
            this.dispStoreyRange.Size = new System.Drawing.Size(224, 29);
            this.dispStoreyRange.TabIndex = 28;
            this.dispStoreyRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispStoreyRange.WordWrap = false;
            // 
            // calAWL
            // 
            this.calAWL.ForeColor = System.Drawing.SystemColors.WindowText;
            this.calAWL.Location = new System.Drawing.Point(9, 205);
            this.calAWL.Margin = new System.Windows.Forms.Padding(6);
            this.calAWL.Name = "calAWL";
            this.calAWL.Size = new System.Drawing.Size(229, 46);
            this.calAWL.TabIndex = 31;
            this.calAWL.Text = "Calculate AWL";
            this.calAWL.UseVisualStyleBackColor = true;
            this.calAWL.Click += new System.EventHandler(this.calAWL_Click);
            // 
            // ETABSTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.EtabsTabGroup);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "ETABSTaskPane";
            this.Size = new System.Drawing.Size(550, 1532);
            this.EtabsTabGroup.ResumeLayout(false);
            this.windLoadPage.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl EtabsTabGroup;
        private System.Windows.Forms.TabPage windLoadPage;
        private System.Windows.Forms.Button setStoreyRange;
        private System.Windows.Forms.TextBox dispStoreyRange;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button getJointCoordinates;
        private System.Windows.Forms.Button getStoreyNames;
        private System.Windows.Forms.Button calAWL;
    }
}
