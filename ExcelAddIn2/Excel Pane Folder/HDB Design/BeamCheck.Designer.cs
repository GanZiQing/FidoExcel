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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.setPcAsvOutRange = new System.Windows.Forms.Button();
            this.dispPcAsvOutRange = new System.Windows.Forms.TextBox();
            this.convertToPcAsv = new System.Windows.Forms.Button();
            this.setPcAsvInRange = new System.Windows.Forms.Button();
            this.dispPcAsvInRange = new System.Windows.Forms.TextBox();
            this.setPcAsOutRange = new System.Windows.Forms.Button();
            this.dispPcAsOutRange = new System.Windows.Forms.TextBox();
            this.convertToPcAs = new System.Windows.Forms.Button();
            this.setPcAsInRange = new System.Windows.Forms.Button();
            this.dispPcAsInRange = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.unmergerTabPage.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.unmergerTabPage);
            this.tabControl1.Location = new System.Drawing.Point(4, 4);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(298, 826);
            this.tabControl1.TabIndex = 3;
            // 
            // unmergerTabPage
            // 
            this.unmergerTabPage.Controls.Add(this.groupBox1);
            this.unmergerTabPage.Controls.Add(this.groupBox2);
            this.unmergerTabPage.Location = new System.Drawing.Point(4, 22);
            this.unmergerTabPage.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.unmergerTabPage.Name = "unmergerTabPage";
            this.unmergerTabPage.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.unmergerTabPage.Size = new System.Drawing.Size(290, 800);
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
            this.groupBox2.Location = new System.Drawing.Point(5, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(283, 210);
            this.groupBox2.TabIndex = 38;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Convert Beam Rebar";
            // 
            // setAsvOutRange
            // 
            this.setAsvOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsvOutRange.Location = new System.Drawing.Point(12, 145);
            this.setAsvOutRange.Name = "setAsvOutRange";
            this.setAsvOutRange.Size = new System.Drawing.Size(125, 25);
            this.setAsvOutRange.TabIndex = 47;
            this.setAsvOutRange.Text = "Set Asv Output Range";
            this.setAsvOutRange.UseVisualStyleBackColor = true;
            // 
            // dispAsvOutRange
            // 
            this.dispAsvOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsvOutRange.Location = new System.Drawing.Point(148, 148);
            this.dispAsvOutRange.Name = "dispAsvOutRange";
            this.dispAsvOutRange.Size = new System.Drawing.Size(124, 20);
            this.dispAsvOutRange.TabIndex = 48;
            this.dispAsvOutRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsvOutRange.WordWrap = false;
            // 
            // convertToAsv
            // 
            this.convertToAsv.ForeColor = System.Drawing.SystemColors.WindowText;
            this.convertToAsv.Location = new System.Drawing.Point(73, 176);
            this.convertToAsv.Name = "convertToAsv";
            this.convertToAsv.Size = new System.Drawing.Size(125, 25);
            this.convertToAsv.TabIndex = 46;
            this.convertToAsv.Text = "Convert to Asv";
            this.convertToAsv.UseVisualStyleBackColor = true;
            this.convertToAsv.Click += new System.EventHandler(this.convertToAsv_Click);
            // 
            // setAsvInRange
            // 
            this.setAsvInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsvInRange.Location = new System.Drawing.Point(12, 113);
            this.setAsvInRange.Name = "setAsvInRange";
            this.setAsvInRange.Size = new System.Drawing.Size(125, 25);
            this.setAsvInRange.TabIndex = 44;
            this.setAsvInRange.Text = "Set Asv Input Range";
            this.setAsvInRange.UseVisualStyleBackColor = true;
            // 
            // dispAsvInRange
            // 
            this.dispAsvInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsvInRange.Location = new System.Drawing.Point(148, 117);
            this.dispAsvInRange.Name = "dispAsvInRange";
            this.dispAsvInRange.Size = new System.Drawing.Size(124, 20);
            this.dispAsvInRange.TabIndex = 45;
            this.dispAsvInRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsvInRange.WordWrap = false;
            // 
            // setAsOutRange
            // 
            this.setAsOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsOutRange.Location = new System.Drawing.Point(12, 50);
            this.setAsOutRange.Name = "setAsOutRange";
            this.setAsOutRange.Size = new System.Drawing.Size(125, 25);
            this.setAsOutRange.TabIndex = 42;
            this.setAsOutRange.Text = "Set As Output Range";
            this.setAsOutRange.UseVisualStyleBackColor = true;
            // 
            // dispAsOutRange
            // 
            this.dispAsOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsOutRange.Location = new System.Drawing.Point(148, 54);
            this.dispAsOutRange.Name = "dispAsOutRange";
            this.dispAsOutRange.Size = new System.Drawing.Size(124, 20);
            this.dispAsOutRange.TabIndex = 43;
            this.dispAsOutRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsOutRange.WordWrap = false;
            // 
            // convertToAs
            // 
            this.convertToAs.ForeColor = System.Drawing.SystemColors.WindowText;
            this.convertToAs.Location = new System.Drawing.Point(73, 82);
            this.convertToAs.Name = "convertToAs";
            this.convertToAs.Size = new System.Drawing.Size(125, 25);
            this.convertToAs.TabIndex = 39;
            this.convertToAs.Text = "Convert to As";
            this.convertToAs.UseVisualStyleBackColor = true;
            this.convertToAs.Click += new System.EventHandler(this.convertToAs_Click);
            // 
            // setAsInRange
            // 
            this.setAsInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setAsInRange.Location = new System.Drawing.Point(12, 19);
            this.setAsInRange.Name = "setAsInRange";
            this.setAsInRange.Size = new System.Drawing.Size(125, 25);
            this.setAsInRange.TabIndex = 32;
            this.setAsInRange.Text = "Set As Input Range";
            this.setAsInRange.UseVisualStyleBackColor = true;
            // 
            // dispAsInRange
            // 
            this.dispAsInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAsInRange.Location = new System.Drawing.Point(148, 23);
            this.dispAsInRange.Name = "dispAsInRange";
            this.dispAsInRange.Size = new System.Drawing.Size(124, 20);
            this.dispAsInRange.TabIndex = 33;
            this.dispAsInRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAsInRange.WordWrap = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.setPcAsvOutRange);
            this.groupBox1.Controls.Add(this.dispPcAsvOutRange);
            this.groupBox1.Controls.Add(this.convertToPcAsv);
            this.groupBox1.Controls.Add(this.setPcAsvInRange);
            this.groupBox1.Controls.Add(this.dispPcAsvInRange);
            this.groupBox1.Controls.Add(this.setPcAsOutRange);
            this.groupBox1.Controls.Add(this.dispPcAsOutRange);
            this.groupBox1.Controls.Add(this.convertToPcAs);
            this.groupBox1.Controls.Add(this.setPcAsInRange);
            this.groupBox1.Controls.Add(this.dispPcAsInRange);
            this.groupBox1.Location = new System.Drawing.Point(5, 221);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(283, 210);
            this.groupBox1.TabIndex = 49;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Convert PC Rebar";
            // 
            // setPcAsvOutRange
            // 
            this.setPcAsvOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPcAsvOutRange.Location = new System.Drawing.Point(12, 145);
            this.setPcAsvOutRange.Name = "setPcAsvOutRange";
            this.setPcAsvOutRange.Size = new System.Drawing.Size(125, 25);
            this.setPcAsvOutRange.TabIndex = 47;
            this.setPcAsvOutRange.Text = "Set Asv Output Range";
            this.setPcAsvOutRange.UseVisualStyleBackColor = true;
            // 
            // dispPcAsvOutRange
            // 
            this.dispPcAsvOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPcAsvOutRange.Location = new System.Drawing.Point(148, 148);
            this.dispPcAsvOutRange.Name = "dispPcAsvOutRange";
            this.dispPcAsvOutRange.Size = new System.Drawing.Size(124, 20);
            this.dispPcAsvOutRange.TabIndex = 48;
            this.dispPcAsvOutRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispPcAsvOutRange.WordWrap = false;
            // 
            // convertToPcAsv
            // 
            this.convertToPcAsv.ForeColor = System.Drawing.SystemColors.WindowText;
            this.convertToPcAsv.Location = new System.Drawing.Point(73, 176);
            this.convertToPcAsv.Name = "convertToPcAsv";
            this.convertToPcAsv.Size = new System.Drawing.Size(125, 25);
            this.convertToPcAsv.TabIndex = 46;
            this.convertToPcAsv.Text = "Convert to Asv";
            this.convertToPcAsv.UseVisualStyleBackColor = true;
            this.convertToPcAsv.Click += new System.EventHandler(this.convertToPcAsv_Click);
            // 
            // setPcAsvInRange
            // 
            this.setPcAsvInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPcAsvInRange.Location = new System.Drawing.Point(12, 113);
            this.setPcAsvInRange.Name = "setPcAsvInRange";
            this.setPcAsvInRange.Size = new System.Drawing.Size(125, 25);
            this.setPcAsvInRange.TabIndex = 44;
            this.setPcAsvInRange.Text = "Set Asv Input Range";
            this.setPcAsvInRange.UseVisualStyleBackColor = true;
            // 
            // dispPcAsvInRange
            // 
            this.dispPcAsvInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPcAsvInRange.Location = new System.Drawing.Point(148, 117);
            this.dispPcAsvInRange.Name = "dispPcAsvInRange";
            this.dispPcAsvInRange.Size = new System.Drawing.Size(124, 20);
            this.dispPcAsvInRange.TabIndex = 45;
            this.dispPcAsvInRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispPcAsvInRange.WordWrap = false;
            // 
            // setPcAsOutRange
            // 
            this.setPcAsOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPcAsOutRange.Location = new System.Drawing.Point(12, 50);
            this.setPcAsOutRange.Name = "setPcAsOutRange";
            this.setPcAsOutRange.Size = new System.Drawing.Size(125, 25);
            this.setPcAsOutRange.TabIndex = 42;
            this.setPcAsOutRange.Text = "Set As Output Range";
            this.setPcAsOutRange.UseVisualStyleBackColor = true;
            // 
            // dispPcAsOutRange
            // 
            this.dispPcAsOutRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPcAsOutRange.Location = new System.Drawing.Point(148, 54);
            this.dispPcAsOutRange.Name = "dispPcAsOutRange";
            this.dispPcAsOutRange.Size = new System.Drawing.Size(124, 20);
            this.dispPcAsOutRange.TabIndex = 43;
            this.dispPcAsOutRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispPcAsOutRange.WordWrap = false;
            // 
            // convertToPcAs
            // 
            this.convertToPcAs.ForeColor = System.Drawing.SystemColors.WindowText;
            this.convertToPcAs.Location = new System.Drawing.Point(73, 82);
            this.convertToPcAs.Name = "convertToPcAs";
            this.convertToPcAs.Size = new System.Drawing.Size(125, 25);
            this.convertToPcAs.TabIndex = 39;
            this.convertToPcAs.Text = "Convert to As";
            this.convertToPcAs.UseVisualStyleBackColor = true;
            this.convertToPcAs.Click += new System.EventHandler(this.convertToPcAs_Click);
            // 
            // setPcAsInRange
            // 
            this.setPcAsInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPcAsInRange.Location = new System.Drawing.Point(12, 19);
            this.setPcAsInRange.Name = "setPcAsInRange";
            this.setPcAsInRange.Size = new System.Drawing.Size(125, 25);
            this.setPcAsInRange.TabIndex = 32;
            this.setPcAsInRange.Text = "Set As Input Range";
            this.setPcAsInRange.UseVisualStyleBackColor = true;
            // 
            // dispPcAsInRange
            // 
            this.dispPcAsInRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPcAsInRange.Location = new System.Drawing.Point(148, 23);
            this.dispPcAsInRange.Name = "dispPcAsInRange";
            this.dispPcAsInRange.Size = new System.Drawing.Size(124, 20);
            this.dispPcAsInRange.TabIndex = 33;
            this.dispPcAsInRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispPcAsInRange.WordWrap = false;
            // 
            // BeamCheck
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "BeamCheck";
            this.Size = new System.Drawing.Size(298, 826);
            this.tabControl1.ResumeLayout(false);
            this.unmergerTabPage.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button setPcAsvOutRange;
        private System.Windows.Forms.TextBox dispPcAsvOutRange;
        private System.Windows.Forms.Button convertToPcAsv;
        private System.Windows.Forms.Button setPcAsvInRange;
        private System.Windows.Forms.TextBox dispPcAsvInRange;
        private System.Windows.Forms.Button setPcAsOutRange;
        private System.Windows.Forms.TextBox dispPcAsOutRange;
        private System.Windows.Forms.Button convertToPcAs;
        private System.Windows.Forms.Button setPcAsInRange;
        private System.Windows.Forms.TextBox dispPcAsInRange;
    }
}
