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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.wallDesignPage = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
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
            this.matchWallRebar = new System.Windows.Forms.Button();
            this.setStatusCol = new System.Windows.Forms.Button();
            this.dispStatusCol = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.wallDesignPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.wallDesignPage);
            this.tabControl1.Location = new System.Drawing.Point(4, 4);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(546, 1524);
            this.tabControl1.TabIndex = 1;
            // 
            // wallDesignPage
            // 
            this.wallDesignPage.Controls.Add(this.groupBox1);
            this.wallDesignPage.Controls.Add(this.matchWallRebar);
            this.wallDesignPage.Location = new System.Drawing.Point(4, 33);
            this.wallDesignPage.Margin = new System.Windows.Forms.Padding(4);
            this.wallDesignPage.Name = "wallDesignPage";
            this.wallDesignPage.Padding = new System.Windows.Forms.Padding(4);
            this.wallDesignPage.Size = new System.Drawing.Size(538, 1487);
            this.wallDesignPage.TabIndex = 0;
            this.wallDesignPage.Text = "Beam Design";
            this.wallDesignPage.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.setStatusCol);
            this.groupBox1.Controls.Add(this.dispStatusCol);
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
            this.groupBox1.Size = new System.Drawing.Size(524, 448);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Set Ranges";
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(10, 179);
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
            this.label5.Location = new System.Drawing.Point(10, 26);
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
            this.dispMatchStoreyCol.Location = new System.Drawing.Point(273, 280);
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
            this.setOutputCol.Location = new System.Drawing.Point(24, 333);
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
            this.setPierLabelRange.Location = new System.Drawing.Point(24, 217);
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
            this.dispStoreyTable.Location = new System.Drawing.Point(273, 132);
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
            this.setRebarTable.Location = new System.Drawing.Point(24, 69);
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
            this.dispRebarTable.Location = new System.Drawing.Point(273, 75);
            this.dispRebarTable.Margin = new System.Windows.Forms.Padding(6);
            this.dispRebarTable.Name = "dispRebarTable";
            this.dispRebarTable.Size = new System.Drawing.Size(224, 29);
            this.dispRebarTable.TabIndex = 33;
            this.dispRebarTable.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispRebarTable.WordWrap = false;
            // 
            // matchWallRebar
            // 
            this.matchWallRebar.ForeColor = System.Drawing.SystemColors.WindowText;
            this.matchWallRebar.Location = new System.Drawing.Point(148, 464);
            this.matchWallRebar.Margin = new System.Windows.Forms.Padding(6);
            this.matchWallRebar.Name = "matchWallRebar";
            this.matchWallRebar.Size = new System.Drawing.Size(229, 46);
            this.matchWallRebar.TabIndex = 38;
            this.matchWallRebar.Text = "Match Reinforcement";
            this.matchWallRebar.UseVisualStyleBackColor = true;
            this.matchWallRebar.Click += new System.EventHandler(this.matchWallRebar_Click);
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
            this.dispStatusCol.Location = new System.Drawing.Point(273, 396);
            this.dispStatusCol.Margin = new System.Windows.Forms.Padding(6);
            this.dispStatusCol.Name = "dispStatusCol";
            this.dispStatusCol.Size = new System.Drawing.Size(224, 29);
            this.dispStatusCol.TabIndex = 44;
            this.dispStatusCol.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispStatusCol.WordWrap = false;
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
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
    }
}
