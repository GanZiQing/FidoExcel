namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class PrintToolsPane
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
            this.PrintPage = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.insertPrintWorkbookHeader = new System.Windows.Forms.Button();
            this.setDestFolder = new System.Windows.Forms.Button();
            this.openDestFolder = new System.Windows.Forms.Button();
            this.dispDestFolder = new System.Windows.Forms.TextBox();
            this.printWorkbooks = new System.Windows.Forms.Button();
            this.overwritePrintPath = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.PrintSelSheetsAdvance = new System.Windows.Forms.Button();
            this.getSheetNames = new System.Windows.Forms.Button();
            this.setSheetNames = new System.Windows.Forms.Button();
            this.PrintMultipleGroup = new System.Windows.Forms.GroupBox();
            this.PrintSelSheets = new System.Windows.Forms.Button();
            this.SetSheetsToPrint = new System.Windows.Forms.Button();
            this.PrintSingleGroup = new System.Windows.Forms.GroupBox();
            this.PrintRangeCheck = new System.Windows.Forms.CheckBox();
            this.PrintCurrentSheet = new System.Windows.Forms.Button();
            this.PrintSettingsGroup = new System.Windows.Forms.GroupBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.OpenPrintFolder = new System.Windows.Forms.Button();
            this.DispPrintFolder = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.DispAppRight = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.DispAppLeft = new System.Windows.Forms.TextBox();
            this.ExcelTabControl = new System.Windows.Forms.TabControl();
            this.dirPage = new System.Windows.Forms.TabPage();
            this.openFilesGroup = new System.Windows.Forms.GroupBox();
            this.dispOpenDelay = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.openFilesInOrder = new System.Windows.Forms.Button();
            this.directoryUserControl = new ExcelAddIn2.DirectoryUserControl();
            this.pdfPage = new System.Windows.Forms.TabPage();
            this.addPageNumGroup = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.dispFontSize = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.dispOffsetX = new System.Windows.Forms.TextBox();
            this.dispOffsetY = new System.Windows.Forms.TextBox();
            this.dispAppendName = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.checkOpenOutput = new System.Windows.Forms.CheckBox();
            this.dispSkipPage = new System.Windows.Forms.TextBox();
            this.addPageNum = new System.Windows.Forms.Button();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.dispFirstPageNum = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.mergePdfGroup = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.generateSections = new System.Windows.Forms.Button();
            this.setRefTitlePage = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dispRefTitlePage = new System.Windows.Forms.TextBox();
            this.dispTitleFontSize = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.insertRefHeader = new System.Windows.Forms.Button();
            this.advancedMerge = new System.Windows.Forms.Button();
            this.openPdfOutFolder = new System.Windows.Forms.Button();
            this.setPdfOutFolder = new System.Windows.Forms.Button();
            this.dispPdfOutFolder = new System.Windows.Forms.TextBox();
            this.dispMergeName = new System.Windows.Forms.TextBox();
            this.labelMergeName = new System.Windows.Forms.TextBox();
            this.basicMergePDF = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.createBookmarksCheck = new System.Windows.Forms.CheckBox();
            this.PrintPage.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.PrintMultipleGroup.SuspendLayout();
            this.PrintSingleGroup.SuspendLayout();
            this.PrintSettingsGroup.SuspendLayout();
            this.ExcelTabControl.SuspendLayout();
            this.dirPage.SuspendLayout();
            this.openFilesGroup.SuspendLayout();
            this.pdfPage.SuspendLayout();
            this.addPageNumGroup.SuspendLayout();
            this.panel1.SuspendLayout();
            this.mergePdfGroup.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // PrintPage
            // 
            this.PrintPage.Controls.Add(this.groupBox4);
            this.PrintPage.Controls.Add(this.groupBox3);
            this.PrintPage.Controls.Add(this.PrintMultipleGroup);
            this.PrintPage.Controls.Add(this.PrintSingleGroup);
            this.PrintPage.Controls.Add(this.PrintSettingsGroup);
            this.PrintPage.Location = new System.Drawing.Point(4, 22);
            this.PrintPage.Name = "PrintPage";
            this.PrintPage.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.PrintPage.Size = new System.Drawing.Size(286, 798);
            this.PrintPage.TabIndex = 1;
            this.PrintPage.Text = "Print Tools";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.insertPrintWorkbookHeader);
            this.groupBox4.Controls.Add(this.setDestFolder);
            this.groupBox4.Controls.Add(this.openDestFolder);
            this.groupBox4.Controls.Add(this.dispDestFolder);
            this.groupBox4.Controls.Add(this.printWorkbooks);
            this.groupBox4.Controls.Add(this.overwritePrintPath);
            this.groupBox4.Location = new System.Drawing.Point(7, 412);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(274, 129);
            this.groupBox4.TabIndex = 12;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Print Workbooks";
            // 
            // insertPrintWorkbookHeader
            // 
            this.insertPrintWorkbookHeader.ForeColor = System.Drawing.SystemColors.WindowText;
            this.insertPrintWorkbookHeader.Location = new System.Drawing.Point(139, 94);
            this.insertPrintWorkbookHeader.Name = "insertPrintWorkbookHeader";
            this.insertPrintWorkbookHeader.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.insertPrintWorkbookHeader.Size = new System.Drawing.Size(125, 25);
            this.insertPrintWorkbookHeader.TabIndex = 103;
            this.insertPrintWorkbookHeader.Text = "Insert Header";
            this.insertPrintWorkbookHeader.UseVisualStyleBackColor = true;
            this.insertPrintWorkbookHeader.Click += new System.EventHandler(this.insertPrintWorkbookHeader_Click);
            // 
            // setDestFolder
            // 
            this.setDestFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setDestFolder.Location = new System.Drawing.Point(8, 18);
            this.setDestFolder.Name = "setDestFolder";
            this.setDestFolder.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setDestFolder.Size = new System.Drawing.Size(125, 25);
            this.setDestFolder.TabIndex = 104;
            this.setDestFolder.Text = "Set Dest. Folder";
            this.setDestFolder.UseVisualStyleBackColor = true;
            // 
            // openDestFolder
            // 
            this.openDestFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openDestFolder.Location = new System.Drawing.Point(143, 18);
            this.openDestFolder.Name = "openDestFolder";
            this.openDestFolder.Size = new System.Drawing.Size(125, 25);
            this.openDestFolder.TabIndex = 105;
            this.openDestFolder.Text = "Open Folder";
            this.openDestFolder.UseVisualStyleBackColor = true;
            // 
            // dispDestFolder
            // 
            this.dispDestFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispDestFolder.Location = new System.Drawing.Point(8, 49);
            this.dispDestFolder.MaxLength = 1000;
            this.dispDestFolder.Name = "dispDestFolder";
            this.dispDestFolder.Size = new System.Drawing.Size(261, 20);
            this.dispDestFolder.TabIndex = 106;
            this.dispDestFolder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // printWorkbooks
            // 
            this.printWorkbooks.ForeColor = System.Drawing.SystemColors.WindowText;
            this.printWorkbooks.Location = new System.Drawing.Point(7, 95);
            this.printWorkbooks.Name = "printWorkbooks";
            this.printWorkbooks.Size = new System.Drawing.Size(125, 25);
            this.printWorkbooks.TabIndex = 45;
            this.printWorkbooks.Text = "Print Workbooks";
            this.printWorkbooks.UseVisualStyleBackColor = true;
            this.printWorkbooks.Click += new System.EventHandler(this.printWorkbooks_Click);
            // 
            // overwritePrintPath
            // 
            this.overwritePrintPath.Checked = true;
            this.overwritePrintPath.CheckState = System.Windows.Forms.CheckState.Checked;
            this.overwritePrintPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.overwritePrintPath.Location = new System.Drawing.Point(8, 72);
            this.overwritePrintPath.Name = "overwritePrintPath";
            this.overwritePrintPath.Size = new System.Drawing.Size(131, 17);
            this.overwritePrintPath.TabIndex = 107;
            this.overwritePrintPath.Text = "Print to file path";
            this.overwritePrintPath.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.PrintSelSheetsAdvance);
            this.groupBox3.Controls.Add(this.getSheetNames);
            this.groupBox3.Controls.Add(this.setSheetNames);
            this.groupBox3.Enabled = false;
            this.groupBox3.Location = new System.Drawing.Point(6, 320);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(274, 85);
            this.groupBox3.TabIndex = 9;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Advance Print Sheets (Not Implemented)";
            // 
            // PrintSelSheetsAdvance
            // 
            this.PrintSelSheetsAdvance.ForeColor = System.Drawing.SystemColors.WindowText;
            this.PrintSelSheetsAdvance.Location = new System.Drawing.Point(9, 50);
            this.PrintSelSheetsAdvance.Name = "PrintSelSheetsAdvance";
            this.PrintSelSheetsAdvance.Size = new System.Drawing.Size(125, 25);
            this.PrintSelSheetsAdvance.TabIndex = 45;
            this.PrintSelSheetsAdvance.Text = "Print Selected Sheets";
            this.PrintSelSheetsAdvance.UseVisualStyleBackColor = true;
            this.PrintSelSheetsAdvance.Click += new System.EventHandler(this.PrintSelSheetsAdvance_Click);
            // 
            // getSheetNames
            // 
            this.getSheetNames.ForeColor = System.Drawing.SystemColors.WindowText;
            this.getSheetNames.Location = new System.Drawing.Point(9, 19);
            this.getSheetNames.Name = "getSheetNames";
            this.getSheetNames.Size = new System.Drawing.Size(125, 25);
            this.getSheetNames.TabIndex = 43;
            this.getSheetNames.Text = "Get Sheet Names";
            this.getSheetNames.UseVisualStyleBackColor = true;
            this.getSheetNames.Click += new System.EventHandler(this.getSheetNames_Click);
            // 
            // setSheetNames
            // 
            this.setSheetNames.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setSheetNames.Location = new System.Drawing.Point(142, 19);
            this.setSheetNames.Name = "setSheetNames";
            this.setSheetNames.Size = new System.Drawing.Size(125, 25);
            this.setSheetNames.TabIndex = 44;
            this.setSheetNames.Text = "Set Sheet Names";
            this.setSheetNames.UseVisualStyleBackColor = true;
            this.setSheetNames.Click += new System.EventHandler(this.setSheetNames_Click);
            // 
            // PrintMultipleGroup
            // 
            this.PrintMultipleGroup.Controls.Add(this.PrintSelSheets);
            this.PrintMultipleGroup.Controls.Add(this.SetSheetsToPrint);
            this.PrintMultipleGroup.Location = new System.Drawing.Point(6, 248);
            this.PrintMultipleGroup.Name = "PrintMultipleGroup";
            this.PrintMultipleGroup.Size = new System.Drawing.Size(274, 66);
            this.PrintMultipleGroup.TabIndex = 3;
            this.PrintMultipleGroup.TabStop = false;
            this.PrintMultipleGroup.Text = "Print Multiple Sheet";
            // 
            // PrintSelSheets
            // 
            this.PrintSelSheets.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.PrintSelSheets.ForeColor = System.Drawing.SystemColors.WindowText;
            this.PrintSelSheets.Location = new System.Drawing.Point(142, 24);
            this.PrintSelSheets.Name = "PrintSelSheets";
            this.PrintSelSheets.Size = new System.Drawing.Size(125, 25);
            this.PrintSelSheets.TabIndex = 8;
            this.PrintSelSheets.Text = "Print Set Sheets";
            this.PrintSelSheets.UseVisualStyleBackColor = true;
            this.PrintSelSheets.Click += new System.EventHandler(this.PrintSelSheets_Click);
            // 
            // SetSheetsToPrint
            // 
            this.SetSheetsToPrint.ForeColor = System.Drawing.SystemColors.WindowText;
            this.SetSheetsToPrint.Location = new System.Drawing.Point(9, 24);
            this.SetSheetsToPrint.Name = "SetSheetsToPrint";
            this.SetSheetsToPrint.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.SetSheetsToPrint.Size = new System.Drawing.Size(125, 25);
            this.SetSheetsToPrint.TabIndex = 7;
            this.SetSheetsToPrint.Text = "Set Sheets to Print";
            this.SetSheetsToPrint.UseVisualStyleBackColor = true;
            // 
            // PrintSingleGroup
            // 
            this.PrintSingleGroup.Controls.Add(this.PrintRangeCheck);
            this.PrintSingleGroup.Controls.Add(this.PrintCurrentSheet);
            this.PrintSingleGroup.Location = new System.Drawing.Point(6, 188);
            this.PrintSingleGroup.Name = "PrintSingleGroup";
            this.PrintSingleGroup.Size = new System.Drawing.Size(274, 55);
            this.PrintSingleGroup.TabIndex = 2;
            this.PrintSingleGroup.TabStop = false;
            this.PrintSingleGroup.Text = "Print Single Sheet";
            // 
            // PrintRangeCheck
            // 
            this.PrintRangeCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.PrintRangeCheck.Location = new System.Drawing.Point(142, 17);
            this.PrintRangeCheck.Name = "PrintRangeCheck";
            this.PrintRangeCheck.Size = new System.Drawing.Size(119, 30);
            this.PrintRangeCheck.TabIndex = 6;
            this.PrintRangeCheck.Text = "Print Selected Range Only";
            this.PrintRangeCheck.UseVisualStyleBackColor = true;
            // 
            // PrintCurrentSheet
            // 
            this.PrintCurrentSheet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.PrintCurrentSheet.ForeColor = System.Drawing.SystemColors.WindowText;
            this.PrintCurrentSheet.Location = new System.Drawing.Point(9, 19);
            this.PrintCurrentSheet.Name = "PrintCurrentSheet";
            this.PrintCurrentSheet.Size = new System.Drawing.Size(125, 25);
            this.PrintCurrentSheet.TabIndex = 5;
            this.PrintCurrentSheet.Text = "Print Current Sheet";
            this.PrintCurrentSheet.UseVisualStyleBackColor = true;
            this.PrintCurrentSheet.Click += new System.EventHandler(this.PrintCurrentSheet_Click);
            // 
            // PrintSettingsGroup
            // 
            this.PrintSettingsGroup.Controls.Add(this.textBox2);
            this.PrintSettingsGroup.Controls.Add(this.OpenPrintFolder);
            this.PrintSettingsGroup.Controls.Add(this.DispPrintFolder);
            this.PrintSettingsGroup.Controls.Add(this.label3);
            this.PrintSettingsGroup.Controls.Add(this.DispAppRight);
            this.PrintSettingsGroup.Controls.Add(this.label1);
            this.PrintSettingsGroup.Controls.Add(this.label2);
            this.PrintSettingsGroup.Controls.Add(this.DispAppLeft);
            this.PrintSettingsGroup.Location = new System.Drawing.Point(6, 6);
            this.PrintSettingsGroup.Name = "PrintSettingsGroup";
            this.PrintSettingsGroup.Size = new System.Drawing.Size(274, 176);
            this.PrintSettingsGroup.TabIndex = 1;
            this.PrintSettingsGroup.TabStop = false;
            this.PrintSettingsGroup.Text = "Print Settings";
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Control;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Location = new System.Drawing.Point(9, 83);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(150, 13);
            this.textBox2.TabIndex = 100;
            this.textBox2.TabStop = false;
            this.textBox2.Text = "Append text to PDF name:";
            // 
            // OpenPrintFolder
            // 
            this.OpenPrintFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.OpenPrintFolder.Location = new System.Drawing.Point(9, 47);
            this.OpenPrintFolder.Name = "OpenPrintFolder";
            this.OpenPrintFolder.Size = new System.Drawing.Size(258, 25);
            this.OpenPrintFolder.TabIndex = 2;
            this.OpenPrintFolder.Text = "Open Target Folder";
            this.OpenPrintFolder.UseVisualStyleBackColor = true;
            this.OpenPrintFolder.Click += new System.EventHandler(this.OpenPrintFolder_Click);
            // 
            // DispPrintFolder
            // 
            this.DispPrintFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DispPrintFolder.Location = new System.Drawing.Point(97, 16);
            this.DispPrintFolder.MaximumSize = new System.Drawing.Size(180, 20);
            this.DispPrintFolder.MaxLength = 100;
            this.DispPrintFolder.MinimumSize = new System.Drawing.Size(160, 20);
            this.DispPrintFolder.Name = "DispPrintFolder";
            this.DispPrintFolder.Size = new System.Drawing.Size(170, 20);
            this.DispPrintFolder.TabIndex = 1;
            this.DispPrintFolder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3
            // 
            this.label3.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label3.Location = new System.Drawing.Point(6, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 20);
            this.label3.TabIndex = 26;
            this.label3.Text = "Folder Name";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // DispAppRight
            // 
            this.DispAppRight.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DispAppRight.Location = new System.Drawing.Point(97, 135);
            this.DispAppRight.MaxLength = 100;
            this.DispAppRight.Name = "DispAppRight";
            this.DispAppRight.Size = new System.Drawing.Size(170, 20);
            this.DispAppRight.TabIndex = 4;
            this.DispAppRight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(6, 105);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 20);
            this.label1.TabIndex = 17;
            this.label1.Text = "Appended Left";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label2.Location = new System.Drawing.Point(6, 135);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 20);
            this.label2.TabIndex = 22;
            this.label2.Text = "Append Right";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // DispAppLeft
            // 
            this.DispAppLeft.ForeColor = System.Drawing.SystemColors.WindowText;
            this.DispAppLeft.Location = new System.Drawing.Point(97, 105);
            this.DispAppLeft.MaxLength = 100;
            this.DispAppLeft.Name = "DispAppLeft";
            this.DispAppLeft.Size = new System.Drawing.Size(170, 20);
            this.DispAppLeft.TabIndex = 3;
            this.DispAppLeft.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // ExcelTabControl
            // 
            this.ExcelTabControl.Controls.Add(this.PrintPage);
            this.ExcelTabControl.Controls.Add(this.dirPage);
            this.ExcelTabControl.Controls.Add(this.pdfPage);
            this.ExcelTabControl.Cursor = System.Windows.Forms.Cursors.Default;
            this.ExcelTabControl.Location = new System.Drawing.Point(3, 3);
            this.ExcelTabControl.Name = "ExcelTabControl";
            this.ExcelTabControl.SelectedIndex = 0;
            this.ExcelTabControl.Size = new System.Drawing.Size(294, 824);
            this.ExcelTabControl.TabIndex = 2;
            // 
            // dirPage
            // 
            this.dirPage.BackColor = System.Drawing.SystemColors.Control;
            this.dirPage.Controls.Add(this.openFilesGroup);
            this.dirPage.Controls.Add(this.directoryUserControl);
            this.dirPage.Location = new System.Drawing.Point(4, 22);
            this.dirPage.Name = "dirPage";
            this.dirPage.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.dirPage.Size = new System.Drawing.Size(286, 798);
            this.dirPage.TabIndex = 3;
            this.dirPage.Text = "Directory";
            // 
            // openFilesGroup
            // 
            this.openFilesGroup.Controls.Add(this.dispOpenDelay);
            this.openFilesGroup.Controls.Add(this.textBox13);
            this.openFilesGroup.Controls.Add(this.openFilesInOrder);
            this.openFilesGroup.Location = new System.Drawing.Point(7, 260);
            this.openFilesGroup.Name = "openFilesGroup";
            this.openFilesGroup.Size = new System.Drawing.Size(274, 81);
            this.openFilesGroup.TabIndex = 6;
            this.openFilesGroup.TabStop = false;
            this.openFilesGroup.Text = "Open Files In Order";
            // 
            // dispOpenDelay
            // 
            this.dispOpenDelay.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOpenDelay.Location = new System.Drawing.Point(111, 19);
            this.dispOpenDelay.MaxLength = 100;
            this.dispOpenDelay.Name = "dispOpenDelay";
            this.dispOpenDelay.Size = new System.Drawing.Size(155, 20);
            this.dispOpenDelay.TabIndex = 44;
            this.dispOpenDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox13
            // 
            this.textBox13.BackColor = System.Drawing.SystemColors.Control;
            this.textBox13.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox13.Location = new System.Drawing.Point(5, 22);
            this.textBox13.Name = "textBox13";
            this.textBox13.ReadOnly = true;
            this.textBox13.ShortcutsEnabled = false;
            this.textBox13.Size = new System.Drawing.Size(100, 13);
            this.textBox13.TabIndex = 45;
            this.textBox13.TabStop = false;
            this.textBox13.Text = "Delay (s)";
            // 
            // openFilesInOrder
            // 
            this.openFilesInOrder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openFilesInOrder.Location = new System.Drawing.Point(10, 45);
            this.openFilesInOrder.Name = "openFilesInOrder";
            this.openFilesInOrder.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.openFilesInOrder.Size = new System.Drawing.Size(260, 25);
            this.openFilesInOrder.TabIndex = 6;
            this.openFilesInOrder.Text = "Open Files In Order";
            this.openFilesInOrder.UseVisualStyleBackColor = true;
            // 
            // directoryUserControl
            // 
            this.directoryUserControl.Location = new System.Drawing.Point(5, 5);
            this.directoryUserControl.Margin = new System.Windows.Forms.Padding(1, 1, 1, 1);
            this.directoryUserControl.Name = "directoryUserControl";
            this.directoryUserControl.Size = new System.Drawing.Size(274, 250);
            this.directoryUserControl.TabIndex = 5;
            // 
            // pdfPage
            // 
            this.pdfPage.BackColor = System.Drawing.SystemColors.Control;
            this.pdfPage.Controls.Add(this.addPageNumGroup);
            this.pdfPage.Controls.Add(this.mergePdfGroup);
            this.pdfPage.Location = new System.Drawing.Point(4, 22);
            this.pdfPage.Name = "pdfPage";
            this.pdfPage.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.pdfPage.Size = new System.Drawing.Size(286, 798);
            this.pdfPage.TabIndex = 2;
            this.pdfPage.Text = "PDF";
            // 
            // addPageNumGroup
            // 
            this.addPageNumGroup.Controls.Add(this.panel1);
            this.addPageNumGroup.Controls.Add(this.dispAppendName);
            this.addPageNumGroup.Controls.Add(this.textBox8);
            this.addPageNumGroup.Controls.Add(this.checkOpenOutput);
            this.addPageNumGroup.Controls.Add(this.dispSkipPage);
            this.addPageNumGroup.Controls.Add(this.addPageNum);
            this.addPageNumGroup.Controls.Add(this.textBox6);
            this.addPageNumGroup.Controls.Add(this.dispFirstPageNum);
            this.addPageNumGroup.Controls.Add(this.textBox4);
            this.addPageNumGroup.Location = new System.Drawing.Point(9, 370);
            this.addPageNumGroup.Name = "addPageNumGroup";
            this.addPageNumGroup.Size = new System.Drawing.Size(274, 220);
            this.addPageNumGroup.TabIndex = 2;
            this.addPageNumGroup.TabStop = false;
            this.addPageNumGroup.Text = "Add Page Number";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textBox3);
            this.panel1.Controls.Add(this.dispFontSize);
            this.panel1.Controls.Add(this.textBox12);
            this.panel1.Controls.Add(this.textBox10);
            this.panel1.Controls.Add(this.textBox11);
            this.panel1.Controls.Add(this.dispOffsetX);
            this.panel1.Controls.Add(this.dispOffsetY);
            this.panel1.Location = new System.Drawing.Point(6, 94);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(268, 70);
            this.panel1.TabIndex = 4;
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.SystemColors.Control;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Location = new System.Drawing.Point(0, 51);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.ShortcutsEnabled = false;
            this.textBox3.Size = new System.Drawing.Size(97, 17);
            this.textBox3.TabIndex = 57;
            this.textBox3.TabStop = false;
            this.textBox3.Text = "FontSize";
            // 
            // dispFontSize
            // 
            this.dispFontSize.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFontSize.Location = new System.Drawing.Point(105, 47);
            this.dispFontSize.MaxLength = 100;
            this.dispFontSize.Name = "dispFontSize";
            this.dispFontSize.Size = new System.Drawing.Size(156, 20);
            this.dispFontSize.TabIndex = 3;
            this.dispFontSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox12
            // 
            this.textBox12.BackColor = System.Drawing.SystemColors.Control;
            this.textBox12.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox12.Location = new System.Drawing.Point(211, 3);
            this.textBox12.Name = "textBox12";
            this.textBox12.ReadOnly = true;
            this.textBox12.ShortcutsEnabled = false;
            this.textBox12.Size = new System.Drawing.Size(27, 13);
            this.textBox12.TabIndex = 100;
            this.textBox12.TabStop = false;
            this.textBox12.Text = "Y";
            this.textBox12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox10
            // 
            this.textBox10.BackColor = System.Drawing.SystemColors.Control;
            this.textBox10.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox10.Location = new System.Drawing.Point(0, 25);
            this.textBox10.Multiline = true;
            this.textBox10.Name = "textBox10";
            this.textBox10.ReadOnly = true;
            this.textBox10.ShortcutsEnabled = false;
            this.textBox10.Size = new System.Drawing.Size(97, 17);
            this.textBox10.TabIndex = 53;
            this.textBox10.TabStop = false;
            this.textBox10.Text = "Page Number Offset";
            // 
            // textBox11
            // 
            this.textBox11.BackColor = System.Drawing.SystemColors.Control;
            this.textBox11.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox11.Location = new System.Drawing.Point(128, 3);
            this.textBox11.Name = "textBox11";
            this.textBox11.ReadOnly = true;
            this.textBox11.ShortcutsEnabled = false;
            this.textBox11.Size = new System.Drawing.Size(27, 13);
            this.textBox11.TabIndex = 100;
            this.textBox11.TabStop = false;
            this.textBox11.Text = "X";
            this.textBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispOffsetX
            // 
            this.dispOffsetX.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOffsetX.Location = new System.Drawing.Point(105, 22);
            this.dispOffsetX.MaxLength = 100;
            this.dispOffsetX.Name = "dispOffsetX";
            this.dispOffsetX.Size = new System.Drawing.Size(74, 20);
            this.dispOffsetX.TabIndex = 1;
            this.dispOffsetX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispOffsetY
            // 
            this.dispOffsetY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOffsetY.Location = new System.Drawing.Point(187, 22);
            this.dispOffsetY.MaxLength = 100;
            this.dispOffsetY.Name = "dispOffsetY";
            this.dispOffsetY.Size = new System.Drawing.Size(74, 20);
            this.dispOffsetY.TabIndex = 2;
            this.dispOffsetY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispAppendName
            // 
            this.dispAppendName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAppendName.Location = new System.Drawing.Point(112, 68);
            this.dispAppendName.MaxLength = 100;
            this.dispAppendName.Name = "dispAppendName";
            this.dispAppendName.Size = new System.Drawing.Size(155, 20);
            this.dispAppendName.TabIndex = 3;
            this.dispAppendName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox8
            // 
            this.textBox8.BackColor = System.Drawing.SystemColors.Control;
            this.textBox8.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox8.Location = new System.Drawing.Point(6, 71);
            this.textBox8.Name = "textBox8";
            this.textBox8.ReadOnly = true;
            this.textBox8.ShortcutsEnabled = false;
            this.textBox8.Size = new System.Drawing.Size(100, 13);
            this.textBox8.TabIndex = 47;
            this.textBox8.TabStop = false;
            this.textBox8.Text = "Append File Name";
            // 
            // checkOpenOutput
            // 
            this.checkOpenOutput.Checked = true;
            this.checkOpenOutput.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkOpenOutput.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkOpenOutput.Location = new System.Drawing.Point(6, 199);
            this.checkOpenOutput.Name = "checkOpenOutput";
            this.checkOpenOutput.Size = new System.Drawing.Size(131, 17);
            this.checkOpenOutput.TabIndex = 6;
            this.checkOpenOutput.Text = "Open Output Files";
            this.checkOpenOutput.UseVisualStyleBackColor = true;
            // 
            // dispSkipPage
            // 
            this.dispSkipPage.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispSkipPage.Location = new System.Drawing.Point(112, 42);
            this.dispSkipPage.MaxLength = 100;
            this.dispSkipPage.Name = "dispSkipPage";
            this.dispSkipPage.Size = new System.Drawing.Size(155, 20);
            this.dispSkipPage.TabIndex = 2;
            this.dispSkipPage.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // addPageNum
            // 
            this.addPageNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.addPageNum.Location = new System.Drawing.Point(6, 168);
            this.addPageNum.Name = "addPageNum";
            this.addPageNum.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.addPageNum.Size = new System.Drawing.Size(260, 25);
            this.addPageNum.TabIndex = 5;
            this.addPageNum.Text = "Add Page Number";
            this.addPageNum.UseVisualStyleBackColor = true;
            this.addPageNum.Click += new System.EventHandler(this.addPageNum_Click);
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.SystemColors.Control;
            this.textBox6.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox6.Location = new System.Drawing.Point(6, 45);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.ShortcutsEnabled = false;
            this.textBox6.Size = new System.Drawing.Size(100, 13);
            this.textBox6.TabIndex = 45;
            this.textBox6.TabStop = false;
            this.textBox6.Text = "Ignore First N Pages";
            // 
            // dispFirstPageNum
            // 
            this.dispFirstPageNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFirstPageNum.Location = new System.Drawing.Point(112, 16);
            this.dispFirstPageNum.MaxLength = 100;
            this.dispFirstPageNum.Name = "dispFirstPageNum";
            this.dispFirstPageNum.Size = new System.Drawing.Size(155, 20);
            this.dispFirstPageNum.TabIndex = 1;
            this.dispFirstPageNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.SystemColors.Control;
            this.textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox4.Location = new System.Drawing.Point(6, 19);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.ShortcutsEnabled = false;
            this.textBox4.Size = new System.Drawing.Size(100, 13);
            this.textBox4.TabIndex = 43;
            this.textBox4.TabStop = false;
            this.textBox4.Text = "First Page Number";
            // 
            // mergePdfGroup
            // 
            this.mergePdfGroup.Controls.Add(this.createBookmarksCheck);
            this.mergePdfGroup.Controls.Add(this.groupBox2);
            this.mergePdfGroup.Controls.Add(this.groupBox1);
            this.mergePdfGroup.Controls.Add(this.openPdfOutFolder);
            this.mergePdfGroup.Controls.Add(this.setPdfOutFolder);
            this.mergePdfGroup.Controls.Add(this.dispPdfOutFolder);
            this.mergePdfGroup.Controls.Add(this.dispMergeName);
            this.mergePdfGroup.Controls.Add(this.labelMergeName);
            this.mergePdfGroup.Controls.Add(this.basicMergePDF);
            this.mergePdfGroup.Location = new System.Drawing.Point(6, 6);
            this.mergePdfGroup.Name = "mergePdfGroup";
            this.mergePdfGroup.Size = new System.Drawing.Size(274, 358);
            this.mergePdfGroup.TabIndex = 1;
            this.mergePdfGroup.TabStop = false;
            this.mergePdfGroup.Text = "Merge PDF";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.generateSections);
            this.groupBox2.Controls.Add(this.setRefTitlePage);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.dispRefTitlePage);
            this.groupBox2.Controls.Add(this.dispTitleFontSize);
            this.groupBox2.Controls.Add(this.textBox5);
            this.groupBox2.Location = new System.Drawing.Point(0, 140);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(274, 163);
            this.groupBox2.TabIndex = 42;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Create Section Dividers";
            // 
            // generateSections
            // 
            this.generateSections.ForeColor = System.Drawing.SystemColors.WindowText;
            this.generateSections.Location = new System.Drawing.Point(6, 130);
            this.generateSections.Name = "generateSections";
            this.generateSections.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.generateSections.Size = new System.Drawing.Size(261, 25);
            this.generateSections.TabIndex = 62;
            this.generateSections.Text = "Create Section Title Page";
            this.generateSections.UseVisualStyleBackColor = true;
            this.generateSections.Click += new System.EventHandler(this.generateSections_Click);
            // 
            // setRefTitlePage
            // 
            this.setRefTitlePage.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRefTitlePage.Location = new System.Drawing.Point(6, 19);
            this.setRefTitlePage.Name = "setRefTitlePage";
            this.setRefTitlePage.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setRefTitlePage.Size = new System.Drawing.Size(125, 25);
            this.setRefTitlePage.TabIndex = 60;
            this.setRefTitlePage.Text = "Set Ref Title Page File";
            this.setRefTitlePage.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button2.Location = new System.Drawing.Point(6, 99);
            this.button2.Name = "button2";
            this.button2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.button2.Size = new System.Drawing.Size(125, 25);
            this.button2.TabIndex = 41;
            this.button2.Text = "FontDialog";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.titleFont_Click);
            // 
            // dispRefTitlePage
            // 
            this.dispRefTitlePage.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispRefTitlePage.Location = new System.Drawing.Point(6, 50);
            this.dispRefTitlePage.MaxLength = 1000;
            this.dispRefTitlePage.Name = "dispRefTitlePage";
            this.dispRefTitlePage.Size = new System.Drawing.Size(261, 20);
            this.dispRefTitlePage.TabIndex = 61;
            this.dispRefTitlePage.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispTitleFontSize
            // 
            this.dispTitleFontSize.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispTitleFontSize.Location = new System.Drawing.Point(111, 73);
            this.dispTitleFontSize.MaxLength = 100;
            this.dispTitleFontSize.Name = "dispTitleFontSize";
            this.dispTitleFontSize.Size = new System.Drawing.Size(156, 20);
            this.dispTitleFontSize.TabIndex = 1;
            this.dispTitleFontSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.SystemColors.Control;
            this.textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox5.Location = new System.Drawing.Point(6, 76);
            this.textBox5.Multiline = true;
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.ShortcutsEnabled = false;
            this.textBox5.Size = new System.Drawing.Size(97, 17);
            this.textBox5.TabIndex = 59;
            this.textBox5.TabStop = false;
            this.textBox5.Text = "Title Font Size";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.insertRefHeader);
            this.groupBox1.Controls.Add(this.advancedMerge);
            this.groupBox1.Location = new System.Drawing.Point(0, 306);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(274, 52);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Advance Merge";
            // 
            // insertRefHeader
            // 
            this.insertRefHeader.ForeColor = System.Drawing.SystemColors.WindowText;
            this.insertRefHeader.Location = new System.Drawing.Point(6, 19);
            this.insertRefHeader.Name = "insertRefHeader";
            this.insertRefHeader.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.insertRefHeader.Size = new System.Drawing.Size(125, 25);
            this.insertRefHeader.TabIndex = 2;
            this.insertRefHeader.Text = "Insert Ref. Header";
            this.insertRefHeader.UseVisualStyleBackColor = true;
            this.insertRefHeader.Click += new System.EventHandler(this.insertRefHeader_Click);
            // 
            // advancedMerge
            // 
            this.advancedMerge.ForeColor = System.Drawing.SystemColors.WindowText;
            this.advancedMerge.Location = new System.Drawing.Point(142, 19);
            this.advancedMerge.Name = "advancedMerge";
            this.advancedMerge.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.advancedMerge.Size = new System.Drawing.Size(125, 25);
            this.advancedMerge.TabIndex = 3;
            this.advancedMerge.Text = "Advanced Merge PDF";
            this.advancedMerge.UseVisualStyleBackColor = true;
            this.advancedMerge.Click += new System.EventHandler(this.advancedMerge_Click);
            // 
            // openPdfOutFolder
            // 
            this.openPdfOutFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openPdfOutFolder.Location = new System.Drawing.Point(142, 19);
            this.openPdfOutFolder.Name = "openPdfOutFolder";
            this.openPdfOutFolder.Size = new System.Drawing.Size(125, 25);
            this.openPdfOutFolder.TabIndex = 2;
            this.openPdfOutFolder.Text = "Open Output Folder";
            this.openPdfOutFolder.UseVisualStyleBackColor = true;
            // 
            // setPdfOutFolder
            // 
            this.setPdfOutFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setPdfOutFolder.Location = new System.Drawing.Point(6, 19);
            this.setPdfOutFolder.Name = "setPdfOutFolder";
            this.setPdfOutFolder.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setPdfOutFolder.Size = new System.Drawing.Size(125, 25);
            this.setPdfOutFolder.TabIndex = 1;
            this.setPdfOutFolder.Text = "Set Output Folder";
            this.setPdfOutFolder.UseVisualStyleBackColor = true;
            // 
            // dispPdfOutFolder
            // 
            this.dispPdfOutFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispPdfOutFolder.Location = new System.Drawing.Point(6, 50);
            this.dispPdfOutFolder.MaxLength = 1000;
            this.dispPdfOutFolder.Name = "dispPdfOutFolder";
            this.dispPdfOutFolder.Size = new System.Drawing.Size(261, 20);
            this.dispPdfOutFolder.TabIndex = 3;
            this.dispPdfOutFolder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispMergeName
            // 
            this.dispMergeName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispMergeName.Location = new System.Drawing.Point(112, 76);
            this.dispMergeName.MaxLength = 100;
            this.dispMergeName.Name = "dispMergeName";
            this.dispMergeName.Size = new System.Drawing.Size(155, 20);
            this.dispMergeName.TabIndex = 4;
            this.dispMergeName.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelMergeName
            // 
            this.labelMergeName.BackColor = System.Drawing.SystemColors.Control;
            this.labelMergeName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelMergeName.Location = new System.Drawing.Point(6, 79);
            this.labelMergeName.Name = "labelMergeName";
            this.labelMergeName.ReadOnly = true;
            this.labelMergeName.ShortcutsEnabled = false;
            this.labelMergeName.Size = new System.Drawing.Size(100, 13);
            this.labelMergeName.TabIndex = 40;
            this.labelMergeName.TabStop = false;
            this.labelMergeName.Text = "Output File Name";
            // 
            // basicMergePDF
            // 
            this.basicMergePDF.ForeColor = System.Drawing.SystemColors.WindowText;
            this.basicMergePDF.Location = new System.Drawing.Point(6, 102);
            this.basicMergePDF.Name = "basicMergePDF";
            this.basicMergePDF.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.basicMergePDF.Size = new System.Drawing.Size(125, 25);
            this.basicMergePDF.TabIndex = 5;
            this.basicMergePDF.Text = "Basic Merge PDF";
            this.basicMergePDF.UseVisualStyleBackColor = true;
            this.basicMergePDF.Click += new System.EventHandler(this.basicMergePDF_Click);
            // 
            // createBookmarksCheck
            // 
            this.createBookmarksCheck.Checked = true;
            this.createBookmarksCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.createBookmarksCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.createBookmarksCheck.Location = new System.Drawing.Point(137, 107);
            this.createBookmarksCheck.Name = "createBookmarksCheck";
            this.createBookmarksCheck.Size = new System.Drawing.Size(131, 17);
            this.createBookmarksCheck.TabIndex = 43;
            this.createBookmarksCheck.Text = "Create Bookmarks";
            this.createBookmarksCheck.UseVisualStyleBackColor = true;
            // 
            // PrintToolsPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.Controls.Add(this.ExcelTabControl);
            this.Name = "PrintToolsPane";
            this.Size = new System.Drawing.Size(300, 830);
            this.PrintPage.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.PrintMultipleGroup.ResumeLayout(false);
            this.PrintSingleGroup.ResumeLayout(false);
            this.PrintSettingsGroup.ResumeLayout(false);
            this.PrintSettingsGroup.PerformLayout();
            this.ExcelTabControl.ResumeLayout(false);
            this.dirPage.ResumeLayout(false);
            this.openFilesGroup.ResumeLayout(false);
            this.openFilesGroup.PerformLayout();
            this.pdfPage.ResumeLayout(false);
            this.addPageNumGroup.ResumeLayout(false);
            this.addPageNumGroup.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.mergePdfGroup.ResumeLayout(false);
            this.mergePdfGroup.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage PrintPage;
        private System.Windows.Forms.GroupBox PrintMultipleGroup;
        private System.Windows.Forms.Button PrintSelSheets;
        private System.Windows.Forms.Button SetSheetsToPrint;
        private System.Windows.Forms.GroupBox PrintSingleGroup;
        private System.Windows.Forms.CheckBox PrintRangeCheck;
        private System.Windows.Forms.Button PrintCurrentSheet;
        private System.Windows.Forms.GroupBox PrintSettingsGroup;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button OpenPrintFolder;
        private System.Windows.Forms.TextBox DispPrintFolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox DispAppRight;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox DispAppLeft;
        private System.Windows.Forms.TabControl ExcelTabControl;
        private System.Windows.Forms.TabPage pdfPage;
        private System.Windows.Forms.GroupBox mergePdfGroup;
        private System.Windows.Forms.Button openPdfOutFolder;
        private System.Windows.Forms.Button setPdfOutFolder;
        private System.Windows.Forms.TextBox dispPdfOutFolder;
        private System.Windows.Forms.Button addPageNum;
        private System.Windows.Forms.TextBox dispMergeName;
        private System.Windows.Forms.TextBox labelMergeName;
        private System.Windows.Forms.Button basicMergePDF;
        private System.Windows.Forms.GroupBox addPageNumGroup;
        private System.Windows.Forms.TextBox dispSkipPage;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox dispFirstPageNum;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button advancedMerge;
        private System.Windows.Forms.TextBox dispAppendName;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.CheckBox checkOpenOutput;
        private System.Windows.Forms.TextBox dispOffsetY;
        private System.Windows.Forms.TextBox dispOffsetX;
        private System.Windows.Forms.TextBox textBox12;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox dispFontSize;
        private System.Windows.Forms.Button insertRefHeader;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox dispTitleFontSize;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button setRefTitlePage;
        private System.Windows.Forms.TextBox dispRefTitlePage;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button generateSections;
        private System.Windows.Forms.TabPage dirPage;
        private System.Windows.Forms.Button getSheetNames;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button PrintSelSheetsAdvance;
        private System.Windows.Forms.Button setSheetNames;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button printWorkbooks;
        private System.Windows.Forms.Button insertPrintWorkbookHeader;
        private System.Windows.Forms.Button openDestFolder;
        private System.Windows.Forms.CheckBox overwritePrintPath;
        private System.Windows.Forms.Button setDestFolder;
        private System.Windows.Forms.TextBox dispDestFolder;
        private DirectoryUserControl directoryUserControl;
        private System.Windows.Forms.GroupBox openFilesGroup;
        private System.Windows.Forms.TextBox dispOpenDelay;
        private System.Windows.Forms.TextBox textBox13;
        private System.Windows.Forms.Button openFilesInOrder;
        private System.Windows.Forms.CheckBox createBookmarksCheck;
    }
}
