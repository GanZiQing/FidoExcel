namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class AutoCadPane
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
            this.dispLineOptions = new System.Windows.Forms.ComboBox();
            this.getLineCoords = new System.Windows.Forms.Button();
            this.AutoCadTabControl = new System.Windows.Forms.TabControl();
            this.printXCheck = new System.Windows.Forms.CheckBox();
            this.printYCheck = new System.Windows.Forms.CheckBox();
            this.printZCheck = new System.Windows.Forms.CheckBox();
            this.printEndCheck = new System.Windows.Forms.CheckBox();
            this.printMidCheck = new System.Windows.Forms.CheckBox();
            this.printStartCheck = new System.Windows.Forms.CheckBox();
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
            this.tabPage1.Text = "Coordinates";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.printEndCheck);
            this.groupBox1.Controls.Add(this.printMidCheck);
            this.groupBox1.Controls.Add(this.printStartCheck);
            this.groupBox1.Controls.Add(this.printZCheck);
            this.groupBox1.Controls.Add(this.printYCheck);
            this.groupBox1.Controls.Add(this.printXCheck);
            this.groupBox1.Controls.Add(this.dispLineOptions);
            this.groupBox1.Controls.Add(this.getLineCoords);
            this.groupBox1.Location = new System.Drawing.Point(15, 11);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox1.Size = new System.Drawing.Size(502, 340);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Line Functions";
            // 
            // dispLineOptions
            // 
            this.dispLineOptions.AutoCompleteCustomSource.AddRange(new string[] {
            "1 Start, mid, end",
            "2 Start, end",
            "3 Mid"});
            this.dispLineOptions.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.dispLineOptions.FormattingEnabled = true;
            this.dispLineOptions.Items.AddRange(new object[] {
            "1 Start, mid, end",
            "2 Start, end",
            "3 Mid"});
            this.dispLineOptions.Location = new System.Drawing.Point(266, 43);
            this.dispLineOptions.Margin = new System.Windows.Forms.Padding(6);
            this.dispLineOptions.Name = "dispLineOptions";
            this.dispLineOptions.Size = new System.Drawing.Size(220, 32);
            this.dispLineOptions.TabIndex = 2;
            // 
            // getLineCoords
            // 
            this.getLineCoords.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getLineCoords.Location = new System.Drawing.Point(11, 35);
            this.getLineCoords.Margin = new System.Windows.Forms.Padding(6);
            this.getLineCoords.Name = "getLineCoords";
            this.getLineCoords.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getLineCoords.Size = new System.Drawing.Size(229, 46);
            this.getLineCoords.TabIndex = 1;
            this.getLineCoords.Text = "Get Coordinates";
            this.getLineCoords.UseVisualStyleBackColor = true;
            this.getLineCoords.Click += new System.EventHandler(this.getLineStartEnd_Click);
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
            // printXCheck
            // 
            this.printXCheck.AutoSize = true;
            this.printXCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printXCheck.Location = new System.Drawing.Point(11, 91);
            this.printXCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printXCheck.Name = "printXCheck";
            this.printXCheck.Size = new System.Drawing.Size(52, 29);
            this.printXCheck.TabIndex = 42;
            this.printXCheck.Text = "X";
            this.printXCheck.UseVisualStyleBackColor = true;
            // 
            // printYCheck
            // 
            this.printYCheck.AutoSize = true;
            this.printYCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printYCheck.Location = new System.Drawing.Point(115, 91);
            this.printYCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printYCheck.Name = "printYCheck";
            this.printYCheck.Size = new System.Drawing.Size(51, 29);
            this.printYCheck.TabIndex = 43;
            this.printYCheck.Text = "Y";
            this.printYCheck.UseVisualStyleBackColor = true;
            // 
            // printZCheck
            // 
            this.printZCheck.AutoSize = true;
            this.printZCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printZCheck.Location = new System.Drawing.Point(218, 91);
            this.printZCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printZCheck.Name = "printZCheck";
            this.printZCheck.Size = new System.Drawing.Size(50, 29);
            this.printZCheck.TabIndex = 44;
            this.printZCheck.Text = "Z";
            this.printZCheck.UseVisualStyleBackColor = true;
            // 
            // printEndCheck
            // 
            this.printEndCheck.AutoSize = true;
            this.printEndCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printEndCheck.Location = new System.Drawing.Point(218, 128);
            this.printEndCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printEndCheck.Name = "printEndCheck";
            this.printEndCheck.Size = new System.Drawing.Size(73, 29);
            this.printEndCheck.TabIndex = 47;
            this.printEndCheck.Text = "End";
            this.printEndCheck.UseVisualStyleBackColor = true;
            // 
            // printMidCheck
            // 
            this.printMidCheck.AutoSize = true;
            this.printMidCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printMidCheck.Location = new System.Drawing.Point(115, 128);
            this.printMidCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printMidCheck.Name = "printMidCheck";
            this.printMidCheck.Size = new System.Drawing.Size(70, 29);
            this.printMidCheck.TabIndex = 46;
            this.printMidCheck.Text = "Mid";
            this.printMidCheck.UseVisualStyleBackColor = true;
            // 
            // printStartCheck
            // 
            this.printStartCheck.AutoSize = true;
            this.printStartCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printStartCheck.Location = new System.Drawing.Point(11, 128);
            this.printStartCheck.Margin = new System.Windows.Forms.Padding(4);
            this.printStartCheck.Name = "printStartCheck";
            this.printStartCheck.Size = new System.Drawing.Size(79, 29);
            this.printStartCheck.TabIndex = 45;
            this.printStartCheck.Text = "Start";
            this.printStartCheck.UseVisualStyleBackColor = true;
            // 
            // AutoCadPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.AutoCadTabControl);
            this.Name = "AutoCadPane";
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
        private System.Windows.Forms.Button getLineCoords;
        private System.Windows.Forms.TabControl AutoCadTabControl;
        private System.Windows.Forms.ComboBox dispLineOptions;
        private System.Windows.Forms.CheckBox printZCheck;
        private System.Windows.Forms.CheckBox printYCheck;
        private System.Windows.Forms.CheckBox printXCheck;
        private System.Windows.Forms.CheckBox printEndCheck;
        private System.Windows.Forms.CheckBox printMidCheck;
        private System.Windows.Forms.CheckBox printStartCheck;
    }
}
