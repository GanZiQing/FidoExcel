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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.drawDropPanel = new System.Windows.Forms.Button();
            this.selectErrorJoint = new System.Windows.Forms.Button();
            this.selectBeamLabel = new System.Windows.Forms.Button();
            this.checkWalls = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.removeUNBack = new System.Windows.Forms.Button();
            this.copyFrameLabel = new System.Windows.Forms.Button();
            this.getFloors = new System.Windows.Forms.Button();
            this.dupeUnits = new System.Windows.Forms.Button();
            this.getSelCoord = new System.Windows.Forms.Button();
            this.getGroups = new System.Windows.Forms.Button();
            this.EtabsTabGroup.SuspendLayout();
            this.EtabsPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // EtabsTabGroup
            // 
            this.EtabsTabGroup.Controls.Add(this.EtabsPage1);
            this.EtabsTabGroup.Location = new System.Drawing.Point(3, 3);
            this.EtabsTabGroup.Name = "EtabsTabGroup";
            this.EtabsTabGroup.SelectedIndex = 0;
            this.EtabsTabGroup.Size = new System.Drawing.Size(294, 1035);
            this.EtabsTabGroup.TabIndex = 4;
            // 
            // EtabsPage1
            // 
            this.EtabsPage1.Controls.Add(this.groupBox2);
            this.EtabsPage1.Controls.Add(this.groupBox1);
            this.EtabsPage1.Location = new System.Drawing.Point(4, 22);
            this.EtabsPage1.Name = "EtabsPage1";
            this.EtabsPage1.Padding = new System.Windows.Forms.Padding(3);
            this.EtabsPage1.Size = new System.Drawing.Size(286, 1009);
            this.EtabsPage1.TabIndex = 0;
            this.EtabsPage1.Text = "EtabsPage1";
            this.EtabsPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.drawDropPanel);
            this.groupBox2.Controls.Add(this.selectErrorJoint);
            this.groupBox2.Controls.Add(this.selectBeamLabel);
            this.groupBox2.Controls.Add(this.checkWalls);
            this.groupBox2.Location = new System.Drawing.Point(5, 216);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(280, 175);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Utilities";
            // 
            // button3
            // 
            this.button3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button3.Location = new System.Drawing.Point(6, 142);
            this.button3.Name = "button3";
            this.button3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button3.Size = new System.Drawing.Size(150, 25);
            this.button3.TabIndex = 18;
            this.button3.Text = "Select Error Joints";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // drawDropPanel
            // 
            this.drawDropPanel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.drawDropPanel.Location = new System.Drawing.Point(5, 111);
            this.drawDropPanel.Name = "drawDropPanel";
            this.drawDropPanel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.drawDropPanel.Size = new System.Drawing.Size(150, 25);
            this.drawDropPanel.TabIndex = 16;
            this.drawDropPanel.Text = "Draw Drop Panel";
            this.drawDropPanel.UseVisualStyleBackColor = true;
            this.drawDropPanel.Click += new System.EventHandler(this.drawDropPanel_Click);
            // 
            // selectErrorJoint
            // 
            this.selectErrorJoint.ForeColor = System.Drawing.SystemColors.WindowText;
            this.selectErrorJoint.Location = new System.Drawing.Point(5, 80);
            this.selectErrorJoint.Name = "selectErrorJoint";
            this.selectErrorJoint.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.selectErrorJoint.Size = new System.Drawing.Size(150, 25);
            this.selectErrorJoint.TabIndex = 15;
            this.selectErrorJoint.Text = "Select Error Joints";
            this.selectErrorJoint.UseVisualStyleBackColor = true;
            this.selectErrorJoint.Click += new System.EventHandler(this.selectErrorJoint_Click);
            // 
            // selectBeamLabel
            // 
            this.selectBeamLabel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.selectBeamLabel.Location = new System.Drawing.Point(5, 48);
            this.selectBeamLabel.Name = "selectBeamLabel";
            this.selectBeamLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.selectBeamLabel.Size = new System.Drawing.Size(150, 25);
            this.selectBeamLabel.TabIndex = 14;
            this.selectBeamLabel.Text = "Select Beam By Label";
            this.selectBeamLabel.UseVisualStyleBackColor = true;
            this.selectBeamLabel.Click += new System.EventHandler(this.selectBeamLabel_Click);
            // 
            // checkWalls
            // 
            this.checkWalls.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkWalls.Location = new System.Drawing.Point(5, 17);
            this.checkWalls.Name = "checkWalls";
            this.checkWalls.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkWalls.Size = new System.Drawing.Size(150, 25);
            this.checkWalls.TabIndex = 10;
            this.checkWalls.Text = "Check Walls";
            this.checkWalls.UseVisualStyleBackColor = true;
            this.checkWalls.Click += new System.EventHandler(this.checkWalls_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.removeUNBack);
            this.groupBox1.Controls.Add(this.copyFrameLabel);
            this.groupBox1.Controls.Add(this.getFloors);
            this.groupBox1.Controls.Add(this.dupeUnits);
            this.groupBox1.Controls.Add(this.getSelCoord);
            this.groupBox1.Controls.Add(this.getGroups);
            this.groupBox1.Location = new System.Drawing.Point(5, 5);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(280, 207);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Unit Duplicator";
            // 
            // removeUNBack
            // 
            this.removeUNBack.ForeColor = System.Drawing.SystemColors.WindowText;
            this.removeUNBack.Location = new System.Drawing.Point(7, 173);
            this.removeUNBack.Name = "removeUNBack";
            this.removeUNBack.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.removeUNBack.Size = new System.Drawing.Size(150, 25);
            this.removeUNBack.TabIndex = 19;
            this.removeUNBack.Text = "Remove UN - Back";
            this.removeUNBack.UseVisualStyleBackColor = true;
            this.removeUNBack.Click += new System.EventHandler(this.removeUNBack_Click);
            // 
            // copyFrameLabel
            // 
            this.copyFrameLabel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.copyFrameLabel.Location = new System.Drawing.Point(5, 142);
            this.copyFrameLabel.Name = "copyFrameLabel";
            this.copyFrameLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.copyFrameLabel.Size = new System.Drawing.Size(150, 25);
            this.copyFrameLabel.TabIndex = 18;
            this.copyFrameLabel.Text = "Copy Frame Label ";
            this.copyFrameLabel.UseVisualStyleBackColor = true;
            this.copyFrameLabel.Click += new System.EventHandler(this.copyFrameLabel_Click);
            // 
            // getFloors
            // 
            this.getFloors.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getFloors.Location = new System.Drawing.Point(5, 80);
            this.getFloors.Name = "getFloors";
            this.getFloors.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getFloors.Size = new System.Drawing.Size(150, 25);
            this.getFloors.TabIndex = 15;
            this.getFloors.Text = "Get Floors";
            this.getFloors.UseVisualStyleBackColor = true;
            this.getFloors.Click += new System.EventHandler(this.getFloors_Click);
            // 
            // dupeUnits
            // 
            this.dupeUnits.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dupeUnits.Location = new System.Drawing.Point(5, 111);
            this.dupeUnits.Name = "dupeUnits";
            this.dupeUnits.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.dupeUnits.Size = new System.Drawing.Size(150, 25);
            this.dupeUnits.TabIndex = 17;
            this.dupeUnits.Text = "Duplicate Units";
            this.dupeUnits.UseVisualStyleBackColor = true;
            this.dupeUnits.Click += new System.EventHandler(this.dupeUnits_Click);
            this.dupeUnits.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dupeUnits_RightClick);
            // 
            // getSelCoord
            // 
            this.getSelCoord.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getSelCoord.Location = new System.Drawing.Point(5, 48);
            this.getSelCoord.Name = "getSelCoord";
            this.getSelCoord.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getSelCoord.Size = new System.Drawing.Size(150, 25);
            this.getSelCoord.TabIndex = 14;
            this.getSelCoord.Text = "Get Selected Coordinates";
            this.getSelCoord.UseVisualStyleBackColor = true;
            this.getSelCoord.Click += new System.EventHandler(this.getSelCoord_Click);
            // 
            // getGroups
            // 
            this.getGroups.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getGroups.Location = new System.Drawing.Point(5, 17);
            this.getGroups.Name = "getGroups";
            this.getGroups.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getGroups.Size = new System.Drawing.Size(150, 25);
            this.getGroups.TabIndex = 10;
            this.getGroups.Text = "Get Groups";
            this.getGroups.UseVisualStyleBackColor = true;
            this.getGroups.Click += new System.EventHandler(this.getGroups_Click);
            // 
            // ETABSTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.EtabsTabGroup);
            this.Name = "ETABSTaskPane";
            this.Size = new System.Drawing.Size(300, 830);
            this.EtabsTabGroup.ResumeLayout(false);
            this.EtabsPage1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
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
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button dupeUnits;
        private System.Windows.Forms.Button drawDropPanel;
        private System.Windows.Forms.Button removeUNBack;
        private System.Windows.Forms.Button copyFrameLabel;
    }
}
