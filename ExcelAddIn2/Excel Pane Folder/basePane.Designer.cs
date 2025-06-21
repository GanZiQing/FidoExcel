namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class basePane
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
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.setXXX = new System.Windows.Forms.Button();
            this.AutoCadTabControl = new System.Windows.Forms.TabControl();
            this.dispXXX = new System.Windows.Forms.TextBox();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.AutoCadTabControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(531, 1484);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dispXXX);
            this.groupBox1.Controls.Add(this.setXXX);
            this.groupBox1.Location = new System.Drawing.Point(15, 11);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox1.Size = new System.Drawing.Size(502, 108);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Group";
            // 
            // setXXX
            // 
            this.setXXX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setXXX.Location = new System.Drawing.Point(11, 35);
            this.setXXX.Margin = new System.Windows.Forms.Padding(6);
            this.setXXX.Name = "setXXX";
            this.setXXX.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setXXX.Size = new System.Drawing.Size(229, 46);
            this.setXXX.TabIndex = 1;
            this.setXXX.Text = "Set XXX";
            this.setXXX.UseVisualStyleBackColor = true;
            // 
            // AutoCadTabControl
            // 
            this.AutoCadTabControl.Controls.Add(this.tabPage1);
            this.AutoCadTabControl.Location = new System.Drawing.Point(6, 6);
            this.AutoCadTabControl.Margin = new System.Windows.Forms.Padding(6);
            this.AutoCadTabControl.Name = "AutoCadTabControl";
            this.AutoCadTabControl.SelectedIndex = 0;
            this.AutoCadTabControl.Size = new System.Drawing.Size(539, 1521);
            this.AutoCadTabControl.TabIndex = 0;
            // 
            // dispXXX
            // 
            this.dispXXX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispXXX.Location = new System.Drawing.Point(266, 43);
            this.dispXXX.Margin = new System.Windows.Forms.Padding(6);
            this.dispXXX.MaxLength = 1000;
            this.dispXXX.Name = "dispXXX";
            this.dispXXX.Size = new System.Drawing.Size(220, 29);
            this.dispXXX.TabIndex = 62;
            this.dispXXX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // autoCadPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.AutoCadTabControl);
            this.Name = "autoCadPane";
            this.Size = new System.Drawing.Size(550, 1532);
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.AutoCadTabControl.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button setXXX;
        private System.Windows.Forms.TabControl AutoCadTabControl;
        private System.Windows.Forms.TextBox dispXXX;
    }
}
