namespace ExcelAddIn2.Excel_Pane_Folder.HDB_Design
{
    partial class WallCheckPane
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.unmergerTabPage = new System.Windows.Forms.TabPage();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.setDecomposeRange = new System.Windows.Forms.Button();
            this.dispDecomposeRange = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tabControl1.SuspendLayout();
            this.unmergerTabPage.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.unmergerTabPage);
            this.tabControl1.Location = new System.Drawing.Point(2, 4);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(546, 1525);
            this.tabControl1.TabIndex = 2;
            // 
            // unmergerTabPage
            // 
            this.unmergerTabPage.Controls.Add(this.groupBox5);
            this.unmergerTabPage.Location = new System.Drawing.Point(4, 33);
            this.unmergerTabPage.Name = "unmergerTabPage";
            this.unmergerTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.unmergerTabPage.Size = new System.Drawing.Size(538, 1488);
            this.unmergerTabPage.TabIndex = 1;
            this.unmergerTabPage.Text = "Wall Check";
            this.unmergerTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label6);
            this.groupBox5.Controls.Add(this.setDecomposeRange);
            this.groupBox5.Controls.Add(this.dispDecomposeRange);
            this.groupBox5.Location = new System.Drawing.Point(7, 7);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox5.Size = new System.Drawing.Size(521, 184);
            this.groupBox5.TabIndex = 35;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Match Rebar";
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label6.Location = new System.Drawing.Point(9, 26);
            this.label6.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(480, 37);
            this.label6.TabIndex = 36;
            this.label6.Text = "Reference Tables:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // setDecomposeRange
            // 
            this.setDecomposeRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setDecomposeRange.Location = new System.Drawing.Point(24, 68);
            this.setDecomposeRange.Margin = new System.Windows.Forms.Padding(6);
            this.setDecomposeRange.Name = "setDecomposeRange";
            this.setDecomposeRange.Size = new System.Drawing.Size(229, 46);
            this.setDecomposeRange.TabIndex = 32;
            this.setDecomposeRange.Text = "Set Table";
            this.setDecomposeRange.UseVisualStyleBackColor = true;
            // 
            // dispDecomposeRange
            // 
            this.dispDecomposeRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispDecomposeRange.Location = new System.Drawing.Point(273, 76);
            this.dispDecomposeRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispDecomposeRange.Name = "dispDecomposeRange";
            this.dispDecomposeRange.Size = new System.Drawing.Size(224, 29);
            this.dispDecomposeRange.TabIndex = 33;
            this.dispDecomposeRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispDecomposeRange.WordWrap = false;
            // 
            // WallCheckPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Name = "WallCheckPane";
            this.Size = new System.Drawing.Size(550, 1532);
            this.tabControl1.ResumeLayout(false);
            this.unmergerTabPage.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage unmergerTabPage;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button setDecomposeRange;
        private System.Windows.Forms.TextBox dispDecomposeRange;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}
