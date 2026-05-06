namespace ExcelAddIn2.Excel_Pane_Folder.HDB_Design
{
    partial class BeamCheck
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
            this.unmergerTabPage = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.setAsvOutRange = new System.Windows.Forms.Button();
            this.dispAsvOutRange = new System.Windows.Forms.TextBox();
            this.convertToAsv = new System.Windows.Forms.Button();
            this.setAsvInRange = new System.Windows.Forms.Button();
            this.dispAsvInRange = new System.Windows.Forms.TextBox();
            this.setAsOutRange = new System.Windows.Forms.Button();
            this.dispAsOutRange = new System.Windows.Forms.TextBox();
            this.convertToAs = new System.Windows.Forms.Button();
            this.setAsInRange = new System.Windows.Forms.Button();
            this.dispAsInRange = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.unmergerTabPage.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.unmergerTabPage);
            this.tabControl1.Location = new System.Drawing.Point(8, 8);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(546, 1525);
            this.tabControl1.TabIndex = 3;
            // 
            // unmergerTabPage
            // 
            this.unmergerTabPage.Controls.Add(this.groupBox2);
            this.unmergerTabPage.Location = new System.Drawing.Point(4, 33);
            this.unmergerTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.unmergerTabPage.Name = "unmergerTabPage";
            this.unmergerTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.unmergerTabPage.Size = new System.Drawing.Size(538, 1488);
            this.unmergerTabPage.TabIndex = 1;
            this.unmergerTabPage.Text = "Beam Check";
            this.unmergerTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.setAsvOutRange);
            this.groupBox2.Controls.Add(this.dispAsvOutRange);
            this.groupBox2.Controls.Add(this.convertToAsv);
            this.groupBox2.Controls.Add(this.setAsvInRange);
            this.groupBox2.Controls.Add(this.dispAsvInRange);
            this.groupBox2.Controls.Add(this.setAsOutRange);
            this.groupBox2.Controls.Add(this.dispAsOutRange);
            this.groupBox2.Controls.Add(this.convertToAs);
            this.groupBox2.Controls.Add(this.setAsInRange);
            this.groupBox2.Controls.Add(this.dispAsInRange);
            this.groupBox2.Location = new System.Drawing.Point(9, 9);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox2.Size = new System.Drawing.Size(519, 387);
            this.groupBox2.TabIndex = 38;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Convert Beam Rebar";
            // 
            // setAsvOutRange
            // 
            this.setAsvOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsvOutRange.Location = new System.Drawing.Point(22, 267);
            this.setAsvOutRange.Margin = new System.Windows.Forms.Padding(6);
            this.setAsvOutRange.Name = "setAsvOutRange";
            this.setAsvOutRange.Size = new System.Drawing.Size(229, 46);
            this.setAsvOutRange.TabIndex = 47;
            this.setAsvOutRange.Text = "Set Asv Output Range";
            this.setAsvOutRange.UseVisualStyleBackColor = true;
            // 
            // dispAsvOutRange
            // 
            this.dispAsvOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsvOutRange.Location = new System.Drawing.Point(271, 274);
            this.dispAsvOutRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispAsvOutRange.Name = "dispAsvOutRange";
            this.dispAsvOutRange.Size = new System.Drawing.Size(224, 29);
            this.dispAsvOutRange.TabIndex = 48;
            this.dispAsvOutRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsvOutRange.WordWrap = false;
            // 
            // convertToAsv
            // 
            this.convertToAsv.ForeColor = System.Drawing.SystemColors.WindowText;
            this.convertToAsv.Location = new System.Drawing.Point(133, 325);
            this.convertToAsv.Margin = new System.Windows.Forms.Padding(6);
            this.convertToAsv.Name = "convertToAsv";
            this.convertToAsv.Size = new System.Drawing.Size(229, 46);
            this.convertToAsv.TabIndex = 46;
            this.convertToAsv.Text = "Convert to Asv";
            this.convertToAsv.UseVisualStyleBackColor = true;
            this.convertToAsv.Click += new System.EventHandler(this.convertToAsv_Click);
            // 
            // setAsvInRange
            // 
            this.setAsvInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsvInRange.Location = new System.Drawing.Point(22, 209);
            this.setAsvInRange.Margin = new System.Windows.Forms.Padding(6);
            this.setAsvInRange.Name = "setAsvInRange";
            this.setAsvInRange.Size = new System.Drawing.Size(229, 46);
            this.setAsvInRange.TabIndex = 44;
            this.setAsvInRange.Text = "Set Asv Input Range";
            this.setAsvInRange.UseVisualStyleBackColor = true;
            // 
            // dispAsvInRange
            // 
            this.dispAsvInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsvInRange.Location = new System.Drawing.Point(271, 216);
            this.dispAsvInRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispAsvInRange.Name = "dispAsvInRange";
            this.dispAsvInRange.Size = new System.Drawing.Size(224, 29);
            this.dispAsvInRange.TabIndex = 45;
            this.dispAsvInRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsvInRange.WordWrap = false;
            // 
            // setAsOutRange
            // 
            this.setAsOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsOutRange.Location = new System.Drawing.Point(22, 93);
            this.setAsOutRange.Margin = new System.Windows.Forms.Padding(6);
            this.setAsOutRange.Name = "setAsOutRange";
            this.setAsOutRange.Size = new System.Drawing.Size(229, 46);
            this.setAsOutRange.TabIndex = 42;
            this.setAsOutRange.Text = "Set As Output Range";
            this.setAsOutRange.UseVisualStyleBackColor = true;
            // 
            // dispAsOutRange
            // 
            this.dispAsOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsOutRange.Location = new System.Drawing.Point(271, 100);
            this.dispAsOutRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispAsOutRange.Name = "dispAsOutRange";
            this.dispAsOutRange.Size = new System.Drawing.Size(224, 29);
            this.dispAsOutRange.TabIndex = 43;
            this.dispAsOutRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsOutRange.WordWrap = false;
            // 
            // convertToAs
            // 
            this.convertToAs.ForeColor = System.Drawing.SystemColors.WindowText;
            this.convertToAs.Location = new System.Drawing.Point(133, 151);
            this.convertToAs.Margin = new System.Windows.Forms.Padding(6);
            this.convertToAs.Name = "convertToAs";
            this.convertToAs.Size = new System.Drawing.Size(229, 46);
            this.convertToAs.TabIndex = 39;
            this.convertToAs.Text = "Convert to As";
            this.convertToAs.UseVisualStyleBackColor = true;
            this.convertToAs.Click += new System.EventHandler(this.convertToAs_Click);
            // 
            // setAsInRange
            // 
            this.setAsInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsInRange.Location = new System.Drawing.Point(22, 35);
            this.setAsInRange.Margin = new System.Windows.Forms.Padding(6);
            this.setAsInRange.Name = "setAsInRange";
            this.setAsInRange.Size = new System.Drawing.Size(229, 46);
            this.setAsInRange.TabIndex = 32;
            this.setAsInRange.Text = "Set As Input Range";
            this.setAsInRange.UseVisualStyleBackColor = true;
            // 
            // dispAsInRange
            // 
            this.dispAsInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsInRange.Location = new System.Drawing.Point(271, 42);
            this.dispAsInRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispAsInRange.Name = "dispAsInRange";
            this.dispAsInRange.Size = new System.Drawing.Size(224, 29);
            this.dispAsInRange.TabIndex = 33;
            this.dispAsInRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsInRange.WordWrap = false;
            // 
            // BeamCheck
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "BeamCheck";
            this.Size = new System.Drawing.Size(546, 1525);
            this.tabControl1.ResumeLayout(false);
            this.unmergerTabPage.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage unmergerTabPage;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button setAsvOutRange;
        private System.Windows.Forms.TextBox dispAsvOutRange;
        private System.Windows.Forms.Button convertToAsv;
        private System.Windows.Forms.Button setAsvInRange;
        private System.Windows.Forms.TextBox dispAsvInRange;
        private System.Windows.Forms.Button setAsOutRange;
        private System.Windows.Forms.TextBox dispAsOutRange;
        private System.Windows.Forms.Button convertToAs;
        private System.Windows.Forms.Button setAsInRange;
        private System.Windows.Forms.TextBox dispAsInRange;
    }
}
