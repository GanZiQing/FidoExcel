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
            this.setBeamTable = new System.Windows.Forms.Button();
            this.dispBeamTable = new System.Windows.Forms.TextBox();
            this.setOutputColumn = new System.Windows.Forms.Button();
            this.dispOutputColumn = new System.Windows.Forms.TextBox();
            this.setShearTable = new System.Windows.Forms.Button();
            this.dispShearTable = new System.Windows.Forms.TextBox();
            this.decomposeTable = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.beamDesignPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.beamDesignPage);
            this.tabControl1.Location = new System.Drawing.Point(3, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(547, 1154);
            this.tabControl1.TabIndex = 0;
            // 
            // beamDesignPage
            // 
            this.beamDesignPage.Controls.Add(this.groupBox1);
            this.beamDesignPage.Location = new System.Drawing.Point(4, 33);
            this.beamDesignPage.Name = "beamDesignPage";
            this.beamDesignPage.Padding = new System.Windows.Forms.Padding(3);
            this.beamDesignPage.Size = new System.Drawing.Size(539, 1117);
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
            this.groupBox1.Location = new System.Drawing.Point(9, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(524, 379);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Read Beam Table";
            // 
            // setBeamTable
            // 
            this.setBeamTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setBeamTable.Location = new System.Drawing.Point(23, 31);
            this.setBeamTable.Margin = new System.Windows.Forms.Padding(6);
            this.setBeamTable.Name = "setBeamTable";
            this.setBeamTable.Size = new System.Drawing.Size(229, 46);
            this.setBeamTable.TabIndex = 32;
            this.setBeamTable.Text = "Set Beam Table";
            this.setBeamTable.UseVisualStyleBackColor = true;
            // 
            // dispBeamTable
            // 
            this.dispBeamTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispBeamTable.Location = new System.Drawing.Point(274, 37);
            this.dispBeamTable.Margin = new System.Windows.Forms.Padding(6);
            this.dispBeamTable.Name = "dispBeamTable";
            this.dispBeamTable.Size = new System.Drawing.Size(224, 29);
            this.dispBeamTable.TabIndex = 33;
            this.dispBeamTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispBeamTable.WordWrap = false;
            // 
            // setOutputColumn
            // 
            this.setOutputColumn.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setOutputColumn.Location = new System.Drawing.Point(23, 89);
            this.setOutputColumn.Margin = new System.Windows.Forms.Padding(6);
            this.setOutputColumn.Name = "setOutputColumn";
            this.setOutputColumn.Size = new System.Drawing.Size(229, 46);
            this.setOutputColumn.TabIndex = 34;
            this.setOutputColumn.Text = "Set Output Column";
            this.setOutputColumn.UseVisualStyleBackColor = true;
            // 
            // dispOutputColumn
            // 
            this.dispOutputColumn.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOutputColumn.Location = new System.Drawing.Point(274, 95);
            this.dispOutputColumn.Margin = new System.Windows.Forms.Padding(6);
            this.dispOutputColumn.Name = "dispOutputColumn";
            this.dispOutputColumn.Size = new System.Drawing.Size(224, 29);
            this.dispOutputColumn.TabIndex = 35;
            this.dispOutputColumn.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispOutputColumn.WordWrap = false;
            // 
            // setShearTable
            // 
            this.setShearTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setShearTable.Location = new System.Drawing.Point(23, 147);
            this.setShearTable.Margin = new System.Windows.Forms.Padding(6);
            this.setShearTable.Name = "setShearTable";
            this.setShearTable.Size = new System.Drawing.Size(229, 46);
            this.setShearTable.TabIndex = 36;
            this.setShearTable.Text = "Set Shear Table";
            this.setShearTable.UseVisualStyleBackColor = true;
            // 
            // dispShearTable
            // 
            this.dispShearTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispShearTable.Location = new System.Drawing.Point(274, 153);
            this.dispShearTable.Margin = new System.Windows.Forms.Padding(6);
            this.dispShearTable.Name = "dispShearTable";
            this.dispShearTable.Size = new System.Drawing.Size(224, 29);
            this.dispShearTable.TabIndex = 37;
            this.dispShearTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispShearTable.WordWrap = false;
            // 
            // decomposeTable
            // 
            this.decomposeTable.ForeColor = System.Drawing.SystemColors.WindowText;
            this.decomposeTable.Location = new System.Drawing.Point(150, 205);
            this.decomposeTable.Margin = new System.Windows.Forms.Padding(6);
            this.decomposeTable.Name = "decomposeTable";
            this.decomposeTable.Size = new System.Drawing.Size(229, 46);
            this.decomposeTable.TabIndex = 38;
            this.decomposeTable.Text = "Decompose Table";
            this.decomposeTable.UseVisualStyleBackColor = true;
            this.decomposeTable.Click += new System.EventHandler(this.decomposeTable_Click);
            // 
            // BeamDesign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "BeamDesign";
            this.Size = new System.Drawing.Size(550, 1532);
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
