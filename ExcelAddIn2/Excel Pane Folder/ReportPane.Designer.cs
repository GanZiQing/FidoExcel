namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class ReportPane
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
            this.directoryUserControl1 = new ExcelAddIn2.DirectoryUserControl();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.importToPpt = new System.Windows.Forms.Button();
            this.setImportRange = new System.Windows.Forms.Button();
            this.dispImportRange = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.getBounds = new System.Windows.Forms.Button();
            this.insertImageBox = new System.Windows.Forms.Button();
            this.dispHeightY = new System.Windows.Forms.TextBox();
            this.dispWidthX = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.dispInsertY = new System.Windows.Forms.TextBox();
            this.dispInsertX = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.openPpt = new System.Windows.Forms.Button();
            this.setPptFile = new System.Windows.Forms.Button();
            this.dispPptFile = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.openSCFolder = new System.Windows.Forms.Button();
            this.setSCFolder = new System.Windows.Forms.Button();
            this.dispSCFolder = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dispLoadDelay = new System.Windows.Forms.TextBox();
            this.dispStartDelay = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.saveEtabsImage = new System.Windows.Forms.Button();
            this.setFloorRange = new System.Windows.Forms.Button();
            this.dispFloorRange = new System.Windows.Forms.TextBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.setScreenshotBounds = new System.Windows.Forms.Button();
            this.getScreenshotBounds = new System.Windows.Forms.Button();
            this.dispScreenshotHeight = new System.Windows.Forms.TextBox();
            this.dispScreenshotWidth = new System.Windows.Forms.TextBox();
            this.testScreenshot = new System.Windows.Forms.Button();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.dispScreenshotY = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.dispScreenshotX = new System.Windows.Forms.TextBox();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.launchScreenshotApp = new System.Windows.Forms.Button();
            this.tabPage1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.directoryUserControl1);
            this.tabPage1.Controls.Add(this.groupBox4);
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(6);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(6);
            this.tabPage1.Size = new System.Drawing.Size(531, 1484);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Report Gen";
            // 
            // directoryUserControl1
            // 
            this.directoryUserControl1.Location = new System.Drawing.Point(15, 9);
            this.directoryUserControl1.Margin = new System.Windows.Forms.Padding(4);
            this.directoryUserControl1.Name = "directoryUserControl1";
            this.directoryUserControl1.Size = new System.Drawing.Size(502, 431);
            this.directoryUserControl1.TabIndex = 9;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.importToPpt);
            this.groupBox4.Controls.Add(this.setImportRange);
            this.groupBox4.Controls.Add(this.dispImportRange);
            this.groupBox4.Location = new System.Drawing.Point(15, 833);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox4.Size = new System.Drawing.Size(502, 149);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Import to Ppt";
            // 
            // importToPpt
            // 
            this.importToPpt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importToPpt.Location = new System.Drawing.Point(11, 92);
            this.importToPpt.Margin = new System.Windows.Forms.Padding(6);
            this.importToPpt.Name = "importToPpt";
            this.importToPpt.Size = new System.Drawing.Size(478, 46);
            this.importToPpt.TabIndex = 27;
            this.importToPpt.Text = "Import To Ppt";
            this.importToPpt.UseVisualStyleBackColor = true;
            this.importToPpt.Click += new System.EventHandler(this.importToPpt_Click);
            // 
            // setImportRange
            // 
            this.setImportRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setImportRange.Location = new System.Drawing.Point(11, 35);
            this.setImportRange.Margin = new System.Windows.Forms.Padding(6);
            this.setImportRange.Name = "setImportRange";
            this.setImportRange.Size = new System.Drawing.Size(229, 46);
            this.setImportRange.TabIndex = 25;
            this.setImportRange.Text = "Set Import Range";
            this.setImportRange.UseVisualStyleBackColor = true;
            // 
            // dispImportRange
            // 
            this.dispImportRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispImportRange.Location = new System.Drawing.Point(262, 41);
            this.dispImportRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispImportRange.Name = "dispImportRange";
            this.dispImportRange.Size = new System.Drawing.Size(224, 29);
            this.dispImportRange.TabIndex = 26;
            this.dispImportRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispImportRange.WordWrap = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.getBounds);
            this.groupBox2.Controls.Add(this.insertImageBox);
            this.groupBox2.Controls.Add(this.dispHeightY);
            this.groupBox2.Controls.Add(this.dispWidthX);
            this.groupBox2.Controls.Add(this.textBox10);
            this.groupBox2.Controls.Add(this.textBox6);
            this.groupBox2.Controls.Add(this.textBox4);
            this.groupBox2.Controls.Add(this.dispInsertY);
            this.groupBox2.Controls.Add(this.dispInsertX);
            this.groupBox2.Controls.Add(this.textBox3);
            this.groupBox2.Location = new System.Drawing.Point(15, 602);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox2.Size = new System.Drawing.Size(502, 219);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Image Position";
            // 
            // getBounds
            // 
            this.getBounds.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getBounds.Location = new System.Drawing.Point(266, 166);
            this.getBounds.Margin = new System.Windows.Forms.Padding(6);
            this.getBounds.Name = "getBounds";
            this.getBounds.Size = new System.Drawing.Size(224, 46);
            this.getBounds.TabIndex = 49;
            this.getBounds.Text = "Get Bounds";
            this.getBounds.UseVisualStyleBackColor = true;
            this.getBounds.Click += new System.EventHandler(this.getBounds_Click);
            // 
            // insertImageBox
            // 
            this.insertImageBox.ForeColor = System.Drawing.SystemColors.WindowText;
            this.insertImageBox.Location = new System.Drawing.Point(16, 166);
            this.insertImageBox.Margin = new System.Windows.Forms.Padding(6);
            this.insertImageBox.Name = "insertImageBox";
            this.insertImageBox.Size = new System.Drawing.Size(224, 46);
            this.insertImageBox.TabIndex = 48;
            this.insertImageBox.Text = "Insert Image Box";
            this.insertImageBox.UseVisualStyleBackColor = true;
            this.insertImageBox.Click += new System.EventHandler(this.insertImageBox_Click);
            // 
            // dispHeightY
            // 
            this.dispHeightY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispHeightY.Location = new System.Drawing.Point(359, 118);
            this.dispHeightY.Margin = new System.Windows.Forms.Padding(6);
            this.dispHeightY.MaxLength = 100;
            this.dispHeightY.Name = "dispHeightY";
            this.dispHeightY.Size = new System.Drawing.Size(125, 29);
            this.dispHeightY.TabIndex = 47;
            this.dispHeightY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispWidthX
            // 
            this.dispWidthX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispWidthX.Location = new System.Drawing.Point(205, 118);
            this.dispWidthX.Margin = new System.Windows.Forms.Padding(6);
            this.dispWidthX.MaxLength = 100;
            this.dispWidthX.Name = "dispWidthX";
            this.dispWidthX.Size = new System.Drawing.Size(125, 29);
            this.dispWidthX.TabIndex = 45;
            this.dispWidthX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox10
            // 
            this.textBox10.BackColor = System.Drawing.SystemColors.Control;
            this.textBox10.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox10.Location = new System.Drawing.Point(11, 124);
            this.textBox10.Margin = new System.Windows.Forms.Padding(6);
            this.textBox10.Name = "textBox10";
            this.textBox10.ReadOnly = true;
            this.textBox10.Size = new System.Drawing.Size(183, 22);
            this.textBox10.TabIndex = 46;
            this.textBox10.TabStop = false;
            this.textBox10.Text = "Dimensions";
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.SystemColors.Control;
            this.textBox6.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox6.Location = new System.Drawing.Point(359, 35);
            this.textBox6.Margin = new System.Windows.Forms.Padding(6);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.Size = new System.Drawing.Size(128, 22);
            this.textBox6.TabIndex = 44;
            this.textBox6.TabStop = false;
            this.textBox6.Text = "Y";
            this.textBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.SystemColors.Control;
            this.textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox4.Location = new System.Drawing.Point(205, 35);
            this.textBox4.Margin = new System.Windows.Forms.Padding(6);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(128, 22);
            this.textBox4.TabIndex = 43;
            this.textBox4.TabStop = false;
            this.textBox4.Text = "X";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispInsertY
            // 
            this.dispInsertY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispInsertY.Location = new System.Drawing.Point(359, 70);
            this.dispInsertY.Margin = new System.Windows.Forms.Padding(6);
            this.dispInsertY.MaxLength = 100;
            this.dispInsertY.Name = "dispInsertY";
            this.dispInsertY.Size = new System.Drawing.Size(125, 29);
            this.dispInsertY.TabIndex = 42;
            this.dispInsertY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispInsertX
            // 
            this.dispInsertX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispInsertX.Location = new System.Drawing.Point(205, 70);
            this.dispInsertX.Margin = new System.Windows.Forms.Padding(6);
            this.dispInsertX.MaxLength = 100;
            this.dispInsertX.Name = "dispInsertX";
            this.dispInsertX.Size = new System.Drawing.Size(125, 29);
            this.dispInsertX.TabIndex = 39;
            this.dispInsertX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.SystemColors.Control;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Location = new System.Drawing.Point(11, 76);
            this.textBox3.Margin = new System.Windows.Forms.Padding(6);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(183, 22);
            this.textBox3.TabIndex = 40;
            this.textBox3.TabStop = false;
            this.textBox3.Text = "Insert Point";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.openPpt);
            this.groupBox1.Controls.Add(this.setPptFile);
            this.groupBox1.Controls.Add(this.dispPptFile);
            this.groupBox1.Location = new System.Drawing.Point(15, 450);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox1.Size = new System.Drawing.Size(502, 140);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Ppt File Def.";
            // 
            // openPpt
            // 
            this.openPpt.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openPpt.Location = new System.Drawing.Point(262, 35);
            this.openPpt.Margin = new System.Windows.Forms.Padding(6);
            this.openPpt.Name = "openPpt";
            this.openPpt.Size = new System.Drawing.Size(229, 46);
            this.openPpt.TabIndex = 14;
            this.openPpt.Text = "Open File";
            this.openPpt.UseVisualStyleBackColor = true;
            // 
            // setPptFile
            // 
            this.setPptFile.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPptFile.Location = new System.Drawing.Point(11, 35);
            this.setPptFile.Margin = new System.Windows.Forms.Padding(6);
            this.setPptFile.Name = "setPptFile";
            this.setPptFile.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setPptFile.Size = new System.Drawing.Size(229, 46);
            this.setPptFile.TabIndex = 12;
            this.setPptFile.Text = "Set Ppt File";
            this.setPptFile.UseVisualStyleBackColor = true;
            // 
            // dispPptFile
            // 
            this.dispPptFile.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPptFile.Location = new System.Drawing.Point(11, 92);
            this.dispPptFile.Margin = new System.Windows.Forms.Padding(6);
            this.dispPptFile.MaxLength = 1000;
            this.dispPptFile.Name = "dispPptFile";
            this.dispPptFile.Size = new System.Drawing.Size(475, 29);
            this.dispPptFile.TabIndex = 13;
            this.dispPptFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(6, 6);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(6);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(539, 1521);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage2.Controls.Add(this.groupBox5);
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Controls.Add(this.groupBox6);
            this.tabPage2.Controls.Add(this.launchScreenshotApp);
            this.tabPage2.Location = new System.Drawing.Point(4, 33);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(6);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(6);
            this.tabPage2.Size = new System.Drawing.Size(531, 1484);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "ETABS Screenshots";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.openSCFolder);
            this.groupBox5.Controls.Add(this.setSCFolder);
            this.groupBox5.Controls.Add(this.dispSCFolder);
            this.groupBox5.Location = new System.Drawing.Point(11, 11);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox5.Size = new System.Drawing.Size(502, 148);
            this.groupBox5.TabIndex = 50;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Screenshot Directory";
            // 
            // openSCFolder
            // 
            this.openSCFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openSCFolder.Location = new System.Drawing.Point(264, 35);
            this.openSCFolder.Margin = new System.Windows.Forms.Padding(6);
            this.openSCFolder.Name = "openSCFolder";
            this.openSCFolder.Size = new System.Drawing.Size(229, 46);
            this.openSCFolder.TabIndex = 48;
            this.openSCFolder.Text = "Open Folder";
            this.openSCFolder.UseVisualStyleBackColor = true;
            // 
            // setSCFolder
            // 
            this.setSCFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSCFolder.Location = new System.Drawing.Point(15, 35);
            this.setSCFolder.Margin = new System.Windows.Forms.Padding(6);
            this.setSCFolder.Name = "setSCFolder";
            this.setSCFolder.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setSCFolder.Size = new System.Drawing.Size(229, 46);
            this.setSCFolder.TabIndex = 47;
            this.setSCFolder.Text = "Set Folder";
            this.setSCFolder.UseVisualStyleBackColor = true;
            // 
            // dispSCFolder
            // 
            this.dispSCFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispSCFolder.Location = new System.Drawing.Point(15, 92);
            this.dispSCFolder.Margin = new System.Windows.Forms.Padding(6);
            this.dispSCFolder.MaxLength = 1000;
            this.dispSCFolder.Name = "dispSCFolder";
            this.dispSCFolder.Size = new System.Drawing.Size(475, 29);
            this.dispSCFolder.TabIndex = 49;
            this.dispSCFolder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dispLoadDelay);
            this.groupBox3.Controls.Add(this.dispStartDelay);
            this.groupBox3.Controls.Add(this.textBox12);
            this.groupBox3.Controls.Add(this.textBox8);
            this.groupBox3.Controls.Add(this.saveEtabsImage);
            this.groupBox3.Controls.Add(this.setFloorRange);
            this.groupBox3.Controls.Add(this.dispFloorRange);
            this.groupBox3.Location = new System.Drawing.Point(11, 471);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox3.Size = new System.Drawing.Size(502, 249);
            this.groupBox3.TabIndex = 33;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "ETABS Print All Floors";
            // 
            // dispLoadDelay
            // 
            this.dispLoadDelay.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispLoadDelay.Location = new System.Drawing.Point(262, 137);
            this.dispLoadDelay.Margin = new System.Windows.Forms.Padding(6);
            this.dispLoadDelay.MaxLength = 100;
            this.dispLoadDelay.Name = "dispLoadDelay";
            this.dispLoadDelay.Size = new System.Drawing.Size(224, 29);
            this.dispLoadDelay.TabIndex = 43;
            this.dispLoadDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispStartDelay
            // 
            this.dispStartDelay.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispStartDelay.Location = new System.Drawing.Point(262, 89);
            this.dispStartDelay.Margin = new System.Windows.Forms.Padding(6);
            this.dispStartDelay.MaxLength = 100;
            this.dispStartDelay.Name = "dispStartDelay";
            this.dispStartDelay.Size = new System.Drawing.Size(224, 29);
            this.dispStartDelay.TabIndex = 45;
            this.dispStartDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox12
            // 
            this.textBox12.BackColor = System.Drawing.SystemColors.Control;
            this.textBox12.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox12.Location = new System.Drawing.Point(16, 142);
            this.textBox12.Margin = new System.Windows.Forms.Padding(6);
            this.textBox12.Name = "textBox12";
            this.textBox12.ReadOnly = true;
            this.textBox12.Size = new System.Drawing.Size(229, 22);
            this.textBox12.TabIndex = 44;
            this.textBox12.TabStop = false;
            this.textBox12.Text = "Load Delay (s)";
            // 
            // textBox8
            // 
            this.textBox8.BackColor = System.Drawing.SystemColors.Control;
            this.textBox8.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox8.Location = new System.Drawing.Point(16, 94);
            this.textBox8.Margin = new System.Windows.Forms.Padding(6);
            this.textBox8.Name = "textBox8";
            this.textBox8.ReadOnly = true;
            this.textBox8.Size = new System.Drawing.Size(229, 22);
            this.textBox8.TabIndex = 46;
            this.textBox8.TabStop = false;
            this.textBox8.Text = "Start Delay (s)";
            // 
            // saveEtabsImage
            // 
            this.saveEtabsImage.ForeColor = System.Drawing.SystemColors.WindowText;
            this.saveEtabsImage.Location = new System.Drawing.Point(11, 185);
            this.saveEtabsImage.Margin = new System.Windows.Forms.Padding(6);
            this.saveEtabsImage.Name = "saveEtabsImage";
            this.saveEtabsImage.Size = new System.Drawing.Size(478, 46);
            this.saveEtabsImage.TabIndex = 30;
            this.saveEtabsImage.Text = "Save Images";
            this.saveEtabsImage.UseVisualStyleBackColor = true;
            this.saveEtabsImage.Click += new System.EventHandler(this.saveEtabsImage_Click);
            // 
            // setFloorRange
            // 
            this.setFloorRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setFloorRange.Location = new System.Drawing.Point(11, 35);
            this.setFloorRange.Margin = new System.Windows.Forms.Padding(6);
            this.setFloorRange.Name = "setFloorRange";
            this.setFloorRange.Size = new System.Drawing.Size(229, 46);
            this.setFloorRange.TabIndex = 28;
            this.setFloorRange.Text = "Set Floor Range";
            this.setFloorRange.UseVisualStyleBackColor = true;
            // 
            // dispFloorRange
            // 
            this.dispFloorRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFloorRange.Location = new System.Drawing.Point(262, 41);
            this.dispFloorRange.Margin = new System.Windows.Forms.Padding(6);
            this.dispFloorRange.Name = "dispFloorRange";
            this.dispFloorRange.Size = new System.Drawing.Size(224, 29);
            this.dispFloorRange.TabIndex = 29;
            this.dispFloorRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispFloorRange.WordWrap = false;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.setScreenshotBounds);
            this.groupBox6.Controls.Add(this.getScreenshotBounds);
            this.groupBox6.Controls.Add(this.dispScreenshotHeight);
            this.groupBox6.Controls.Add(this.dispScreenshotWidth);
            this.groupBox6.Controls.Add(this.testScreenshot);
            this.groupBox6.Controls.Add(this.textBox14);
            this.groupBox6.Controls.Add(this.dispScreenshotY);
            this.groupBox6.Controls.Add(this.textBox15);
            this.groupBox6.Controls.Add(this.dispScreenshotX);
            this.groupBox6.Controls.Add(this.textBox18);
            this.groupBox6.Controls.Add(this.textBox13);
            this.groupBox6.Location = new System.Drawing.Point(11, 170);
            this.groupBox6.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox6.Size = new System.Drawing.Size(502, 290);
            this.groupBox6.TabIndex = 14;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Screenshot Position";
            // 
            // setScreenshotBounds
            // 
            this.setScreenshotBounds.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setScreenshotBounds.Location = new System.Drawing.Point(11, 166);
            this.setScreenshotBounds.Margin = new System.Windows.Forms.Padding(6);
            this.setScreenshotBounds.Name = "setScreenshotBounds";
            this.setScreenshotBounds.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setScreenshotBounds.Size = new System.Drawing.Size(224, 46);
            this.setScreenshotBounds.TabIndex = 14;
            this.setScreenshotBounds.Text = "Show Form";
            this.setScreenshotBounds.UseVisualStyleBackColor = true;
            this.setScreenshotBounds.Click += new System.EventHandler(this.setScreenshotBounds_Click);
            // 
            // getScreenshotBounds
            // 
            this.getScreenshotBounds.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getScreenshotBounds.Location = new System.Drawing.Point(268, 166);
            this.getScreenshotBounds.Margin = new System.Windows.Forms.Padding(6);
            this.getScreenshotBounds.Name = "getScreenshotBounds";
            this.getScreenshotBounds.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.getScreenshotBounds.Size = new System.Drawing.Size(224, 46);
            this.getScreenshotBounds.TabIndex = 15;
            this.getScreenshotBounds.Text = "Get Bounds";
            this.getScreenshotBounds.UseVisualStyleBackColor = true;
            this.getScreenshotBounds.Click += new System.EventHandler(this.getScreenshotBounds_Click);
            // 
            // dispScreenshotHeight
            // 
            this.dispScreenshotHeight.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispScreenshotHeight.Location = new System.Drawing.Point(361, 70);
            this.dispScreenshotHeight.Margin = new System.Windows.Forms.Padding(6);
            this.dispScreenshotHeight.MaxLength = 100;
            this.dispScreenshotHeight.Name = "dispScreenshotHeight";
            this.dispScreenshotHeight.Size = new System.Drawing.Size(125, 29);
            this.dispScreenshotHeight.TabIndex = 47;
            this.dispScreenshotHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispScreenshotWidth
            // 
            this.dispScreenshotWidth.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispScreenshotWidth.Location = new System.Drawing.Point(207, 70);
            this.dispScreenshotWidth.Margin = new System.Windows.Forms.Padding(6);
            this.dispScreenshotWidth.MaxLength = 100;
            this.dispScreenshotWidth.Name = "dispScreenshotWidth";
            this.dispScreenshotWidth.Size = new System.Drawing.Size(125, 29);
            this.dispScreenshotWidth.TabIndex = 45;
            this.dispScreenshotWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // testScreenshot
            // 
            this.testScreenshot.ForeColor = System.Drawing.SystemColors.WindowText;
            this.testScreenshot.Location = new System.Drawing.Point(11, 223);
            this.testScreenshot.Margin = new System.Windows.Forms.Padding(6);
            this.testScreenshot.Name = "testScreenshot";
            this.testScreenshot.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.testScreenshot.Size = new System.Drawing.Size(480, 46);
            this.testScreenshot.TabIndex = 12;
            this.testScreenshot.Text = "Test Screenshot";
            this.testScreenshot.UseVisualStyleBackColor = true;
            this.testScreenshot.Click += new System.EventHandler(this.testScreenshot_Click);
            // 
            // textBox14
            // 
            this.textBox14.BackColor = System.Drawing.SystemColors.Control;
            this.textBox14.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox14.Location = new System.Drawing.Point(359, 35);
            this.textBox14.Margin = new System.Windows.Forms.Padding(6);
            this.textBox14.Name = "textBox14";
            this.textBox14.ReadOnly = true;
            this.textBox14.Size = new System.Drawing.Size(128, 22);
            this.textBox14.TabIndex = 44;
            this.textBox14.TabStop = false;
            this.textBox14.Text = "Y";
            this.textBox14.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispScreenshotY
            // 
            this.dispScreenshotY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispScreenshotY.Location = new System.Drawing.Point(361, 118);
            this.dispScreenshotY.Margin = new System.Windows.Forms.Padding(6);
            this.dispScreenshotY.MaxLength = 100;
            this.dispScreenshotY.Name = "dispScreenshotY";
            this.dispScreenshotY.Size = new System.Drawing.Size(125, 29);
            this.dispScreenshotY.TabIndex = 42;
            this.dispScreenshotY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox15
            // 
            this.textBox15.BackColor = System.Drawing.SystemColors.Control;
            this.textBox15.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox15.Location = new System.Drawing.Point(205, 35);
            this.textBox15.Margin = new System.Windows.Forms.Padding(6);
            this.textBox15.Name = "textBox15";
            this.textBox15.ReadOnly = true;
            this.textBox15.Size = new System.Drawing.Size(128, 22);
            this.textBox15.TabIndex = 43;
            this.textBox15.TabStop = false;
            this.textBox15.Text = "X";
            this.textBox15.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispScreenshotX
            // 
            this.dispScreenshotX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispScreenshotX.Location = new System.Drawing.Point(207, 118);
            this.dispScreenshotX.Margin = new System.Windows.Forms.Padding(6);
            this.dispScreenshotX.MaxLength = 100;
            this.dispScreenshotX.Name = "dispScreenshotX";
            this.dispScreenshotX.Size = new System.Drawing.Size(125, 29);
            this.dispScreenshotX.TabIndex = 39;
            this.dispScreenshotX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox18
            // 
            this.textBox18.BackColor = System.Drawing.SystemColors.Control;
            this.textBox18.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox18.Location = new System.Drawing.Point(13, 124);
            this.textBox18.Margin = new System.Windows.Forms.Padding(6);
            this.textBox18.Name = "textBox18";
            this.textBox18.ReadOnly = true;
            this.textBox18.Size = new System.Drawing.Size(183, 22);
            this.textBox18.TabIndex = 40;
            this.textBox18.TabStop = false;
            this.textBox18.Text = "Insert Point";
            // 
            // textBox13
            // 
            this.textBox13.BackColor = System.Drawing.SystemColors.Control;
            this.textBox13.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox13.Location = new System.Drawing.Point(13, 76);
            this.textBox13.Margin = new System.Windows.Forms.Padding(6);
            this.textBox13.Name = "textBox13";
            this.textBox13.ReadOnly = true;
            this.textBox13.Size = new System.Drawing.Size(183, 22);
            this.textBox13.TabIndex = 46;
            this.textBox13.TabStop = false;
            this.textBox13.Text = "Dimensions";
            // 
            // launchScreenshotApp
            // 
            this.launchScreenshotApp.ForeColor = System.Drawing.SystemColors.WindowText;
            this.launchScreenshotApp.Location = new System.Drawing.Point(22, 731);
            this.launchScreenshotApp.Margin = new System.Windows.Forms.Padding(6);
            this.launchScreenshotApp.Name = "launchScreenshotApp";
            this.launchScreenshotApp.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.launchScreenshotApp.Size = new System.Drawing.Size(477, 46);
            this.launchScreenshotApp.TabIndex = 13;
            this.launchScreenshotApp.Text = "Launch Screenshot App";
            this.launchScreenshotApp.UseVisualStyleBackColor = true;
            this.launchScreenshotApp.Click += new System.EventHandler(this.launchScreenshotApp_Click);
            // 
            // ReportPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "ReportPane";
            this.Size = new System.Drawing.Size(550, 1532);
            this.tabPage1.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button setPptFile;
        private System.Windows.Forms.TextBox dispPptFile;
        private System.Windows.Forms.Button openPpt;
        private System.Windows.Forms.TextBox dispImportRange;
        private System.Windows.Forms.Button setImportRange;
        private System.Windows.Forms.Button importToPpt;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox dispInsertX;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Button insertImageBox;
        private System.Windows.Forms.TextBox dispHeightY;
        private System.Windows.Forms.TextBox dispWidthX;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox dispInsertY;
        private System.Windows.Forms.Button getBounds;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button testScreenshot;
        private System.Windows.Forms.Button launchScreenshotApp;
        private System.Windows.Forms.Button setScreenshotBounds;
        private System.Windows.Forms.Button getScreenshotBounds;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.TextBox dispScreenshotHeight;
        private System.Windows.Forms.TextBox dispScreenshotWidth;
        private System.Windows.Forms.TextBox textBox13;
        private System.Windows.Forms.TextBox textBox14;
        private System.Windows.Forms.TextBox textBox15;
        private System.Windows.Forms.TextBox dispScreenshotY;
        private System.Windows.Forms.TextBox dispScreenshotX;
        private System.Windows.Forms.TextBox textBox18;
        private System.Windows.Forms.TextBox dispStartDelay;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox dispLoadDelay;
        private System.Windows.Forms.TextBox textBox12;
        private System.Windows.Forms.Button saveEtabsImage;
        private System.Windows.Forms.Button setFloorRange;
        private System.Windows.Forms.TextBox dispFloorRange;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button openSCFolder;
        private System.Windows.Forms.Button setSCFolder;
        private System.Windows.Forms.TextBox dispSCFolder;
        private DirectoryUserControl directoryUserControl1;
    }
}
