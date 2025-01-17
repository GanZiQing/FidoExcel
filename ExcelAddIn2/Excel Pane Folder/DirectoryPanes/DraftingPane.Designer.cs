namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class DraftingPane
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
            this.ExcelTabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.textBox19 = new System.Windows.Forms.TextBox();
            this.dispFontSizeSheetNum = new System.Windows.Forms.TextBox();
            this.dispTotalDwgNum = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.dispTotalSheetX = new System.Windows.Forms.TextBox();
            this.dispThisSheetY = new System.Windows.Forms.TextBox();
            this.dispTotalSheetY = new System.Windows.Forms.TextBox();
            this.dispThisSheetX = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox17 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.editFilesSheetNum = new System.Windows.Forms.Button();
            this.renameFilesCheck = new System.Windows.Forms.CheckBox();
            this.addSheetNumberCheck = new System.Windows.Forms.CheckBox();
            this.getFileInfo = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.ExcelTabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // ExcelTabControl
            // 
            this.ExcelTabControl.Controls.Add(this.tabPage1);
            this.ExcelTabControl.Cursor = System.Windows.Forms.Cursors.Default;
            this.ExcelTabControl.Location = new System.Drawing.Point(6, 6);
            this.ExcelTabControl.Margin = new System.Windows.Forms.Padding(6);
            this.ExcelTabControl.Name = "ExcelTabControl";
            this.ExcelTabControl.SelectedIndex = 0;
            this.ExcelTabControl.Size = new System.Drawing.Size(539, 1521);
            this.ExcelTabControl.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.groupBox6);
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(531, 1484);
            this.tabPage1.TabIndex = 4;
            this.tabPage1.Text = "Drafting";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.textBox18);
            this.groupBox6.Controls.Add(this.textBox19);
            this.groupBox6.Controls.Add(this.dispFontSizeSheetNum);
            this.groupBox6.Controls.Add(this.dispTotalDwgNum);
            this.groupBox6.Controls.Add(this.textBox15);
            this.groupBox6.Controls.Add(this.dispTotalSheetX);
            this.groupBox6.Controls.Add(this.dispThisSheetY);
            this.groupBox6.Controls.Add(this.dispTotalSheetY);
            this.groupBox6.Controls.Add(this.dispThisSheetX);
            this.groupBox6.Controls.Add(this.textBox9);
            this.groupBox6.Controls.Add(this.textBox17);
            this.groupBox6.Controls.Add(this.textBox16);
            this.groupBox6.Location = new System.Drawing.Point(15, 6);
            this.groupBox6.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox6.Size = new System.Drawing.Size(502, 233);
            this.groupBox6.TabIndex = 47;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Settings";
            // 
            // textBox18
            // 
            this.textBox18.BackColor = System.Drawing.SystemColors.Control;
            this.textBox18.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox18.Location = new System.Drawing.Point(22, 185);
            this.textBox18.Margin = new System.Windows.Forms.Padding(6);
            this.textBox18.Multiline = true;
            this.textBox18.Name = "textBox18";
            this.textBox18.ReadOnly = true;
            this.textBox18.ShortcutsEnabled = false;
            this.textBox18.Size = new System.Drawing.Size(178, 31);
            this.textBox18.TabIndex = 118;
            this.textBox18.TabStop = false;
            this.textBox18.Text = "Font Size";
            // 
            // textBox19
            // 
            this.textBox19.BackColor = System.Drawing.SystemColors.Control;
            this.textBox19.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox19.Location = new System.Drawing.Point(22, 144);
            this.textBox19.Margin = new System.Windows.Forms.Padding(6);
            this.textBox19.Multiline = true;
            this.textBox19.Name = "textBox19";
            this.textBox19.ReadOnly = true;
            this.textBox19.ShortcutsEnabled = false;
            this.textBox19.Size = new System.Drawing.Size(178, 31);
            this.textBox19.TabIndex = 120;
            this.textBox19.TabStop = false;
            this.textBox19.Text = "Total DWG Number";
            // 
            // dispFontSizeSheetNum
            // 
            this.dispFontSizeSheetNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFontSizeSheetNum.Location = new System.Drawing.Point(211, 187);
            this.dispFontSizeSheetNum.Margin = new System.Windows.Forms.Padding(6);
            this.dispFontSizeSheetNum.MaxLength = 100;
            this.dispFontSizeSheetNum.Name = "dispFontSizeSheetNum";
            this.dispFontSizeSheetNum.Size = new System.Drawing.Size(132, 29);
            this.dispFontSizeSheetNum.TabIndex = 117;
            this.dispFontSizeSheetNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispTotalDwgNum
            // 
            this.dispTotalDwgNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispTotalDwgNum.Location = new System.Drawing.Point(211, 146);
            this.dispTotalDwgNum.Margin = new System.Windows.Forms.Padding(6);
            this.dispTotalDwgNum.MaxLength = 100;
            this.dispTotalDwgNum.Name = "dispTotalDwgNum";
            this.dispTotalDwgNum.Size = new System.Drawing.Size(132, 29);
            this.dispTotalDwgNum.TabIndex = 119;
            this.dispTotalDwgNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox15
            // 
            this.textBox15.BackColor = System.Drawing.SystemColors.Control;
            this.textBox15.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox15.Location = new System.Drawing.Point(378, 31);
            this.textBox15.Margin = new System.Windows.Forms.Padding(6);
            this.textBox15.Name = "textBox15";
            this.textBox15.ReadOnly = true;
            this.textBox15.ShortcutsEnabled = false;
            this.textBox15.Size = new System.Drawing.Size(101, 22);
            this.textBox15.TabIndex = 100;
            this.textBox15.TabStop = false;
            this.textBox15.Text = "Y Coord.";
            this.textBox15.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispTotalSheetX
            // 
            this.dispTotalSheetX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispTotalSheetX.Location = new System.Drawing.Point(211, 105);
            this.dispTotalSheetX.Margin = new System.Windows.Forms.Padding(6);
            this.dispTotalSheetX.MaxLength = 100;
            this.dispTotalSheetX.Name = "dispTotalSheetX";
            this.dispTotalSheetX.Size = new System.Drawing.Size(132, 29);
            this.dispTotalSheetX.TabIndex = 101;
            this.dispTotalSheetX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispThisSheetY
            // 
            this.dispThisSheetY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispThisSheetY.Location = new System.Drawing.Point(361, 65);
            this.dispThisSheetY.Margin = new System.Windows.Forms.Padding(6);
            this.dispThisSheetY.MaxLength = 100;
            this.dispThisSheetY.Name = "dispThisSheetY";
            this.dispThisSheetY.Size = new System.Drawing.Size(132, 29);
            this.dispThisSheetY.TabIndex = 2;
            this.dispThisSheetY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispTotalSheetY
            // 
            this.dispTotalSheetY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispTotalSheetY.Location = new System.Drawing.Point(361, 105);
            this.dispTotalSheetY.Margin = new System.Windows.Forms.Padding(6);
            this.dispTotalSheetY.MaxLength = 100;
            this.dispTotalSheetY.Name = "dispTotalSheetY";
            this.dispTotalSheetY.Size = new System.Drawing.Size(132, 29);
            this.dispTotalSheetY.TabIndex = 102;
            this.dispTotalSheetY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispThisSheetX
            // 
            this.dispThisSheetX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispThisSheetX.Location = new System.Drawing.Point(211, 65);
            this.dispThisSheetX.Margin = new System.Windows.Forms.Padding(6);
            this.dispThisSheetX.MaxLength = 100;
            this.dispThisSheetX.Name = "dispThisSheetX";
            this.dispThisSheetX.Size = new System.Drawing.Size(132, 29);
            this.dispThisSheetX.TabIndex = 1;
            this.dispThisSheetX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox9
            // 
            this.textBox9.BackColor = System.Drawing.SystemColors.Control;
            this.textBox9.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox9.Location = new System.Drawing.Point(20, 105);
            this.textBox9.Margin = new System.Windows.Forms.Padding(6);
            this.textBox9.Multiline = true;
            this.textBox9.Name = "textBox9";
            this.textBox9.ReadOnly = true;
            this.textBox9.ShortcutsEnabled = false;
            this.textBox9.Size = new System.Drawing.Size(178, 31);
            this.textBox9.TabIndex = 57;
            this.textBox9.TabStop = false;
            this.textBox9.Text = "Total Sheet No.";
            // 
            // textBox17
            // 
            this.textBox17.BackColor = System.Drawing.SystemColors.Control;
            this.textBox17.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox17.Location = new System.Drawing.Point(227, 31);
            this.textBox17.Margin = new System.Windows.Forms.Padding(6);
            this.textBox17.Name = "textBox17";
            this.textBox17.ReadOnly = true;
            this.textBox17.ShortcutsEnabled = false;
            this.textBox17.Size = new System.Drawing.Size(101, 22);
            this.textBox17.TabIndex = 100;
            this.textBox17.TabStop = false;
            this.textBox17.Text = "X Coord.";
            this.textBox17.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox16
            // 
            this.textBox16.BackColor = System.Drawing.SystemColors.Control;
            this.textBox16.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox16.Location = new System.Drawing.Point(20, 63);
            this.textBox16.Margin = new System.Windows.Forms.Padding(6);
            this.textBox16.Multiline = true;
            this.textBox16.Name = "textBox16";
            this.textBox16.ReadOnly = true;
            this.textBox16.ShortcutsEnabled = false;
            this.textBox16.Size = new System.Drawing.Size(178, 31);
            this.textBox16.TabIndex = 53;
            this.textBox16.TabStop = false;
            this.textBox16.Text = "This Sheet No.";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.editFilesSheetNum);
            this.groupBox3.Controls.Add(this.renameFilesCheck);
            this.groupBox3.Controls.Add(this.addSheetNumberCheck);
            this.groupBox3.Controls.Add(this.getFileInfo);
            this.groupBox3.Location = new System.Drawing.Point(15, 249);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox3.Size = new System.Drawing.Size(502, 225);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Add Sheet Numbers";
            // 
            // editFilesSheetNum
            // 
            this.editFilesSheetNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.editFilesSheetNum.Location = new System.Drawing.Point(13, 92);
            this.editFilesSheetNum.Margin = new System.Windows.Forms.Padding(6);
            this.editFilesSheetNum.Name = "editFilesSheetNum";
            this.editFilesSheetNum.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.editFilesSheetNum.Size = new System.Drawing.Size(477, 46);
            this.editFilesSheetNum.TabIndex = 112;
            this.editFilesSheetNum.Text = "Edit Files";
            this.editFilesSheetNum.UseVisualStyleBackColor = true;
            this.editFilesSheetNum.Click += new System.EventHandler(this.editFilesSheetNum_Click);
            // 
            // renameFilesCheck
            // 
            this.renameFilesCheck.Checked = true;
            this.renameFilesCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.renameFilesCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.renameFilesCheck.Location = new System.Drawing.Point(20, 151);
            this.renameFilesCheck.Margin = new System.Windows.Forms.Padding(6);
            this.renameFilesCheck.Name = "renameFilesCheck";
            this.renameFilesCheck.Size = new System.Drawing.Size(240, 31);
            this.renameFilesCheck.TabIndex = 111;
            this.renameFilesCheck.Text = "Rename File";
            this.renameFilesCheck.UseVisualStyleBackColor = true;
            // 
            // addSheetNumberCheck
            // 
            this.addSheetNumberCheck.Checked = true;
            this.addSheetNumberCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.addSheetNumberCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.addSheetNumberCheck.Location = new System.Drawing.Point(20, 185);
            this.addSheetNumberCheck.Margin = new System.Windows.Forms.Padding(6);
            this.addSheetNumberCheck.Name = "addSheetNumberCheck";
            this.addSheetNumberCheck.Size = new System.Drawing.Size(301, 31);
            this.addSheetNumberCheck.TabIndex = 110;
            this.addSheetNumberCheck.Text = "Add Sheet Numbers to PDF";
            this.addSheetNumberCheck.UseVisualStyleBackColor = true;
            // 
            // getFileInfo
            // 
            this.getFileInfo.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getFileInfo.Location = new System.Drawing.Point(13, 33);
            this.getFileInfo.Margin = new System.Windows.Forms.Padding(6);
            this.getFileInfo.Name = "getFileInfo";
            this.getFileInfo.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getFileInfo.Size = new System.Drawing.Size(477, 46);
            this.getFileInfo.TabIndex = 6;
            this.getFileInfo.Text = "Get File Information";
            this.getFileInfo.UseVisualStyleBackColor = true;
            this.getFileInfo.Click += new System.EventHandler(this.getFileInfo_Click);
            // 
            // Drafting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ExcelTabControl);
            this.Name = "Drafting";
            this.Size = new System.Drawing.Size(550, 1532);
            this.ExcelTabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl ExcelTabControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.TextBox textBox18;
        private System.Windows.Forms.TextBox textBox19;
        private System.Windows.Forms.TextBox dispFontSizeSheetNum;
        private System.Windows.Forms.TextBox dispTotalDwgNum;
        private System.Windows.Forms.TextBox textBox15;
        private System.Windows.Forms.TextBox dispTotalSheetX;
        private System.Windows.Forms.TextBox dispThisSheetY;
        private System.Windows.Forms.TextBox dispTotalSheetY;
        private System.Windows.Forms.TextBox dispThisSheetX;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.TextBox textBox17;
        private System.Windows.Forms.TextBox textBox16;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button editFilesSheetNum;
        private System.Windows.Forms.CheckBox renameFilesCheck;
        private System.Windows.Forms.CheckBox addSheetNumberCheck;
        private System.Windows.Forms.Button getFileInfo;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}
