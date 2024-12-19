namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class BeamDesign
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.beamDesignPage = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.decomposeTable = new System.Windows.Forms.Button();
            this.setShearTable = new System.Windows.Forms.Button();
            this.dispShearTable = new System.Windows.Forms.TextBox();
            this.setOutputColumn = new System.Windows.Forms.Button();
            this.dispOutputColumn = new System.Windows.Forms.TextBox();
            this.setBeamTable = new System.Windows.Forms.Button();
            this.dispBeamTable = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.beamDesignPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.beamDesignPage);
            this.tabControl1.Location = new System.Drawing.Point(2, 2);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(298, 625);
            this.tabControl1.TabIndex = 0;
            // 
            // beamDesignPage
            // 
            this.beamDesignPage.Controls.Add(this.groupBox1);
            this.beamDesignPage.Location = new System.Drawing.Point(4, 22);
            this.beamDesignPage.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.beamDesignPage.Name = "beamDesignPage";
            this.beamDesignPage.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.beamDesignPage.Size = new System.Drawing.Size(290, 599);
            this.beamDesignPage.TabIndex = 0;
            this.beamDesignPage.Text = "Beam Design";
            this.beamDesignPage.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.decomposeTable);
            this.groupBox1.Controls.Add(this.setShearTable);
            this.groupBox1.Controls.Add(this.dispShearTable);
            this.groupBox1.Controls.Add(this.setOutputColumn);
            this.groupBox1.Controls.Add(this.dispOutputColumn);
            this.groupBox1.Controls.Add(this.setBeamTable);
            this.groupBox1.Controls.Add(this.dispBeamTable);
            this.groupBox1.Location = new System.Drawing.Point(5, 3);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(286, 205);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Read Beam Table";
            // 
            // decomposeTable
            // 
            this.decomposeTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.decomposeTable.Location = new System.Drawing.Point(82, 111);
            this.decomposeTable.Name = "decomposeTable";
            this.decomposeTable.Size = new System.Drawing.Size(125, 25);
            this.decomposeTable.TabIndex = 38;
            this.decomposeTable.Text = "Decompose Table";
            this.decomposeTable.UseVisualStyleBackColor = true;
            this.decomposeTable.Click += new System.EventHandler(this.decomposeTable_Click);
            // 
            // setShearTable
            // 
            this.setShearTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setShearTable.Location = new System.Drawing.Point(13, 80);
            this.setShearTable.Name = "setShearTable";
            this.setShearTable.Size = new System.Drawing.Size(125, 25);
            this.setShearTable.TabIndex = 36;
            this.setShearTable.Text = "Set Shear Table";
            this.setShearTable.UseVisualStyleBackColor = true;
            // 
            // dispShearTable
            // 
            this.dispShearTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispShearTable.Location = new System.Drawing.Point(149, 83);
            this.dispShearTable.Name = "dispShearTable";
            this.dispShearTable.Size = new System.Drawing.Size(124, 20);
            this.dispShearTable.TabIndex = 37;
            this.dispShearTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispShearTable.WordWrap = false;
            // 
            // setOutputColumn
            // 
            this.setOutputColumn.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setOutputColumn.Location = new System.Drawing.Point(13, 48);
            this.setOutputColumn.Name = "setOutputColumn";
            this.setOutputColumn.Size = new System.Drawing.Size(125, 25);
            this.setOutputColumn.TabIndex = 34;
            this.setOutputColumn.Text = "Set Output Column";
            this.setOutputColumn.UseVisualStyleBackColor = true;
            // 
            // dispOutputColumn
            // 
            this.dispOutputColumn.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOutputColumn.Location = new System.Drawing.Point(149, 51);
            this.dispOutputColumn.Name = "dispOutputColumn";
            this.dispOutputColumn.Size = new System.Drawing.Size(124, 20);
            this.dispOutputColumn.TabIndex = 35;
            this.dispOutputColumn.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispOutputColumn.WordWrap = false;
            // 
            // setBeamTable
            // 
            this.setBeamTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setBeamTable.Location = new System.Drawing.Point(13, 17);
            this.setBeamTable.Name = "setBeamTable";
            this.setBeamTable.Size = new System.Drawing.Size(125, 25);
            this.setBeamTable.TabIndex = 32;
            this.setBeamTable.Text = "Set Beam Table";
            this.setBeamTable.UseVisualStyleBackColor = true;
            this.setBeamTable.Click += new System.EventHandler(this.setBeamTable_Click);
            // 
            // dispBeamTable
            // 
            this.dispBeamTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispBeamTable.Location = new System.Drawing.Point(149, 20);
            this.dispBeamTable.Name = "dispBeamTable";
            this.dispBeamTable.Size = new System.Drawing.Size(124, 20);
            this.dispBeamTable.TabIndex = 33;
            this.dispBeamTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispBeamTable.WordWrap = false;
            // 
            // BeamDesign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "BeamDesign";
            this.Size = new System.Drawing.Size(300, 830);
            this.tabControl1.ResumeLayout(false);
            this.beamDesignPage.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage beamDesignPage;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button setShearTable;
        private System.Windows.Forms.TextBox dispShearTable;
        private System.Windows.Forms.Button setOutputColumn;
        private System.Windows.Forms.TextBox dispOutputColumn;
        private System.Windows.Forms.Button setBeamTable;
        private System.Windows.Forms.TextBox dispBeamTable;
        private System.Windows.Forms.Button decomposeTable;
    }
}
