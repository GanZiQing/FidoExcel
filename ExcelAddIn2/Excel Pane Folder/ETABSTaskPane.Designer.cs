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
            this.EtabsTabGroup = new System.Windows.Forms.TabControl();
            this.EtabsPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.getGroups = new System.Windows.Forms.Button();
            this.getSelCoord = new System.Windows.Forms.Button();
            this.getFloors = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.selectErrorJoint = new System.Windows.Forms.Button();
            this.selectBeamLabel = new System.Windows.Forms.Button();
            this.checkWalls = new System.Windows.Forms.Button();
            this.EtabsTabGroup.SuspendLayout();
            this.EtabsPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // EtabsTabGroup
            // 
            this.EtabsTabGroup.Controls.Add(this.EtabsPage1);
            this.EtabsTabGroup.Location = new System.Drawing.Point(6, 6);
            this.EtabsTabGroup.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.EtabsTabGroup.Name = "EtabsTabGroup";
            this.EtabsTabGroup.SelectedIndex = 0;
            this.EtabsTabGroup.Size = new System.Drawing.Size(539, 1911);
            this.EtabsTabGroup.TabIndex = 4;
            // 
            // EtabsPage1
            // 
            this.EtabsPage1.Controls.Add(this.groupBox2);
            this.EtabsPage1.Controls.Add(this.groupBox1);
            this.EtabsPage1.Location = new System.Drawing.Point(4, 33);
            this.EtabsPage1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.EtabsPage1.Name = "EtabsPage1";
            this.EtabsPage1.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.EtabsPage1.Size = new System.Drawing.Size(531, 1874);
            this.EtabsPage1.TabIndex = 0;
            this.EtabsPage1.Text = "EtabsPage1";
            this.EtabsPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.getFloors);
            this.groupBox1.Controls.Add(this.getSelCoord);
            this.groupBox1.Controls.Add(this.getGroups);
            this.groupBox1.Location = new System.Drawing.Point(9, 9);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(513, 224);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Unit Duplicator";
            // 
            // getGroups
            // 
            this.getGroups.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getGroups.Location = new System.Drawing.Point(9, 31);
            this.getGroups.Margin = new System.Windows.Forms.Padding(6);
            this.getGroups.Name = "getGroups";
            this.getGroups.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getGroups.Size = new System.Drawing.Size(229, 46);
            this.getGroups.TabIndex = 10;
            this.getGroups.Text = "Get Groups";
            this.getGroups.UseVisualStyleBackColor = true;
            this.getGroups.Click += new System.EventHandler(this.getGroups_Click);
            // 
            // getSelCoord
            // 
            this.getSelCoord.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getSelCoord.Location = new System.Drawing.Point(9, 89);
            this.getSelCoord.Margin = new System.Windows.Forms.Padding(6);
            this.getSelCoord.Name = "getSelCoord";
            this.getSelCoord.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getSelCoord.Size = new System.Drawing.Size(370, 46);
            this.getSelCoord.TabIndex = 14;
            this.getSelCoord.Text = "Get Selected Coordinates";
            this.getSelCoord.UseVisualStyleBackColor = true;
            this.getSelCoord.Click += new System.EventHandler(this.getSelCoord_Click);
            // 
            // getFloors
            // 
            this.getFloors.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getFloors.Location = new System.Drawing.Point(9, 147);
            this.getFloors.Margin = new System.Windows.Forms.Padding(6);
            this.getFloors.Name = "getFloors";
            this.getFloors.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getFloors.Size = new System.Drawing.Size(370, 46);
            this.getFloors.TabIndex = 15;
            this.getFloors.Text = "Get Floors";
            this.getFloors.UseVisualStyleBackColor = true;
            this.getFloors.Click += new System.EventHandler(this.getFloors_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.selectErrorJoint);
            this.groupBox2.Controls.Add(this.selectBeamLabel);
            this.groupBox2.Controls.Add(this.checkWalls);
            this.groupBox2.Location = new System.Drawing.Point(9, 239);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(513, 220);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Utilities";
            // 
            // selectErrorJoint
            // 
            this.selectErrorJoint.ForeColor = System.Drawing.SystemColors.WindowText;
            this.selectErrorJoint.Location = new System.Drawing.Point(9, 147);
            this.selectErrorJoint.Margin = new System.Windows.Forms.Padding(6);
            this.selectErrorJoint.Name = "selectErrorJoint";
            this.selectErrorJoint.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.selectErrorJoint.Size = new System.Drawing.Size(289, 46);
            this.selectErrorJoint.TabIndex = 15;
            this.selectErrorJoint.Text = "Select Error Joints";
            this.selectErrorJoint.UseVisualStyleBackColor = true;
            this.selectErrorJoint.Click += new System.EventHandler(this.selectErrorJoint_Click);
            // 
            // selectBeamLabel
            // 
            this.selectBeamLabel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.selectBeamLabel.Location = new System.Drawing.Point(9, 89);
            this.selectBeamLabel.Margin = new System.Windows.Forms.Padding(6);
            this.selectBeamLabel.Name = "selectBeamLabel";
            this.selectBeamLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.selectBeamLabel.Size = new System.Drawing.Size(289, 46);
            this.selectBeamLabel.TabIndex = 14;
            this.selectBeamLabel.Text = "Select Beam By Label";
            this.selectBeamLabel.UseVisualStyleBackColor = true;
            this.selectBeamLabel.Click += new System.EventHandler(this.selectBeamLabel_Click);
            // 
            // checkWalls
            // 
            this.checkWalls.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkWalls.Location = new System.Drawing.Point(9, 31);
            this.checkWalls.Margin = new System.Windows.Forms.Padding(6);
            this.checkWalls.Name = "checkWalls";
            this.checkWalls.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkWalls.Size = new System.Drawing.Size(229, 46);
            this.checkWalls.TabIndex = 10;
            this.checkWalls.Text = "Check Walls";
            this.checkWalls.UseVisualStyleBackColor = true;
            this.checkWalls.Click += new System.EventHandler(this.checkWalls_Click);
            // 
            // ETABSTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.EtabsTabGroup);
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "ETABSTaskPane";
            this.Size = new System.Drawing.Size(550, 1532);
            this.EtabsTabGroup.ResumeLayout(false);
            this.EtabsPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl EtabsTabGroup;
        private System.Windows.Forms.TabPage EtabsPage1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button getGroups;
        private System.Windows.Forms.Button getSelCoord;
        private System.Windows.Forms.Button getFloors;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button selectErrorJoint;
        private System.Windows.Forms.Button selectBeamLabel;
        private System.Windows.Forms.Button checkWalls;
    }
}
