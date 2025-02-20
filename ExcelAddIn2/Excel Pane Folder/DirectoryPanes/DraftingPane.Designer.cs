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
            this.testCoordinateGroup = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dispIncrement = new System.Windows.Forms.TextBox();
            this.testAddCoordinate = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.dispValidCustomFont = new System.Windows.Forms.TextBox();
            this.setFontFolder = new System.Windows.Forms.Button();
            this.dispFontPath = new System.Windows.Forms.TextBox();
            this.dispFontName = new System.Windows.Forms.ComboBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.dispFontSizeSheetNum = new System.Windows.Forms.TextBox();
            this.dispTotalSheetX = new System.Windows.Forms.TextBox();
            this.dispThisSheetY = new System.Windows.Forms.TextBox();
            this.dispTotalSheetY = new System.Windows.Forms.TextBox();
            this.dispThisSheetX = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox17 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.setOutputFolder = new System.Windows.Forms.Button();
            this.openOutputFolder = new System.Windows.Forms.Button();
            this.dispOutputFolder = new System.Windows.Forms.TextBox();
            this.editFilesSheetNum = new System.Windows.Forms.Button();
            this.textBox19 = new System.Windows.Forms.TextBox();
            this.renameFilesCheck = new System.Windows.Forms.CheckBox();
            this.dispTotalDwgNum = new System.Windows.Forms.TextBox();
            this.addSheetNumberCheck = new System.Windows.Forms.CheckBox();
            this.getFileInfo = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.ExcelTabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.testCoordinateGroup.SuspendLayout();
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
            this.tabPage1.Controls.Add(this.testCoordinateGroup);
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
            // testCoordinateGroup
            // 
            this.testCoordinateGroup.Controls.Add(this.textBox1);
            this.testCoordinateGroup.Controls.Add(this.dispIncrement);
            this.testCoordinateGroup.Controls.Add(this.testAddCoordinate);
            this.testCoordinateGroup.Location = new System.Drawing.Point(15, 6);
            this.testCoordinateGroup.Margin = new System.Windows.Forms.Padding(4);
            this.testCoordinateGroup.Name = "testCoordinateGroup";
            this.testCoordinateGroup.Padding = new System.Windows.Forms.Padding(4);
            this.testCoordinateGroup.Size = new System.Drawing.Size(504, 133);
            this.testCoordinateGroup.TabIndex = 48;
            this.testCoordinateGroup.TabStop = false;
            this.testCoordinateGroup.Text = "Test Coordinates";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(22, 31);
            this.textBox1.Margin = new System.Windows.Forms.Padding(6);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ShortcutsEnabled = false;
            this.textBox1.Size = new System.Drawing.Size(191, 31);
            this.textBox1.TabIndex = 120;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "Coordinate Increment";
            // 
            // dispIncrement
            // 
            this.dispIncrement.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispIncrement.Location = new System.Drawing.Point(226, 33);
            this.dispIncrement.Margin = new System.Windows.Forms.Padding(6);
            this.dispIncrement.MaxLength = 100;
            this.dispIncrement.Name = "dispIncrement";
            this.dispIncrement.Size = new System.Drawing.Size(270, 29);
            this.dispIncrement.TabIndex = 119;
            this.dispIncrement.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // testAddCoordinate
            // 
            this.testAddCoordinate.ForeColor = System.Drawing.SystemColors.WindowText;
            this.testAddCoordinate.Location = new System.Drawing.Point(15, 74);
            this.testAddCoordinate.Margin = new System.Windows.Forms.Padding(6);
            this.testAddCoordinate.Name = "testAddCoordinate";
            this.testAddCoordinate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.testAddCoordinate.Size = new System.Drawing.Size(477, 46);
            this.testAddCoordinate.TabIndex = 113;
            this.testAddCoordinate.Text = "Test Add Coordinate";
            this.testAddCoordinate.UseVisualStyleBackColor = true;
            this.testAddCoordinate.Click += new System.EventHandler(this.testAddCoordinate_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.dispValidCustomFont);
            this.groupBox6.Controls.Add(this.setFontFolder);
            this.groupBox6.Controls.Add(this.dispFontPath);
            this.groupBox6.Controls.Add(this.dispFontName);
            this.groupBox6.Controls.Add(this.textBox2);
            this.groupBox6.Controls.Add(this.textBox3);
            this.groupBox6.Controls.Add(this.textBox18);
            this.groupBox6.Controls.Add(this.textBox15);
            this.groupBox6.Controls.Add(this.dispFontSizeSheetNum);
            this.groupBox6.Controls.Add(this.dispTotalSheetX);
            this.groupBox6.Controls.Add(this.dispThisSheetY);
            this.groupBox6.Controls.Add(this.dispTotalSheetY);
            this.groupBox6.Controls.Add(this.dispThisSheetX);
            this.groupBox6.Controls.Add(this.textBox9);
            this.groupBox6.Controls.Add(this.textBox17);
            this.groupBox6.Controls.Add(this.textBox16);
            this.groupBox6.Location = new System.Drawing.Point(15, 147);
            this.groupBox6.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox6.Size = new System.Drawing.Size(502, 343);
            this.groupBox6.TabIndex = 47;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Settings";
            // 
            // dispValidCustomFont
            // 
            this.dispValidCustomFont.BackColor = System.Drawing.SystemColors.Control;
            this.dispValidCustomFont.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dispValidCustomFont.Location = new System.Drawing.Point(22, 280);
            this.dispValidCustomFont.Margin = new System.Windows.Forms.Padding(6);
            this.dispValidCustomFont.Multiline = true;
            this.dispValidCustomFont.Name = "dispValidCustomFont";
            this.dispValidCustomFont.ReadOnly = true;
            this.dispValidCustomFont.ShortcutsEnabled = false;
            this.dispValidCustomFont.Size = new System.Drawing.Size(474, 46);
            this.dispValidCustomFont.TabIndex = 124;
            this.dispValidCustomFont.TabStop = false;
            this.dispValidCustomFont.Text = "Custom Font Path: Not set";
            // 
            // setFontFolder
            // 
            this.setFontFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setFontFolder.Location = new System.Drawing.Point(15, 222);
            this.setFontFolder.Margin = new System.Windows.Forms.Padding(6);
            this.setFontFolder.Name = "setFontFolder";
            this.setFontFolder.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setFontFolder.Size = new System.Drawing.Size(229, 46);
            this.setFontFolder.TabIndex = 121;
            this.setFontFolder.Text = "Set Custom Font";
            this.setFontFolder.UseVisualStyleBackColor = true;
            // 
            // dispFontPath
            // 
            this.dispFontPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFontPath.Location = new System.Drawing.Point(256, 230);
            this.dispFontPath.Margin = new System.Windows.Forms.Padding(6);
            this.dispFontPath.MaxLength = 1000;
            this.dispFontPath.Name = "dispFontPath";
            this.dispFontPath.Size = new System.Drawing.Size(240, 29);
            this.dispFontPath.TabIndex = 123;
            this.dispFontPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dispFontName
            // 
            this.dispFontName.AutoCompleteCustomSource.AddRange(new string[] {
            "Arial",
            "Times New Roman",
            "Courier New",
            "Verdana",
            "Lucida Console",
            "Symbol",
            "Custom"});
            this.dispFontName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.dispFontName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.dispFontName.FormattingEnabled = true;
            this.dispFontName.Items.AddRange(new object[] {
            "Arial",
            "Times New Roman",
            "Courier New",
            "Verdana",
            "Lucida Console",
            "Symbol",
            "Custom"});
            this.dispFontName.Location = new System.Drawing.Point(358, 181);
            this.dispFontName.Margin = new System.Windows.Forms.Padding(6);
            this.dispFontName.Name = "dispFontName";
            this.dispFontName.Size = new System.Drawing.Size(136, 32);
            this.dispFontName.TabIndex = 49;
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Control;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Location = new System.Drawing.Point(378, 146);
            this.textBox2.Margin = new System.Windows.Forms.Padding(6);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.ShortcutsEnabled = false;
            this.textBox2.Size = new System.Drawing.Size(101, 22);
            this.textBox2.TabIndex = 119;
            this.textBox2.TabStop = false;
            this.textBox2.Text = "Name";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.SystemColors.Control;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Location = new System.Drawing.Point(227, 146);
            this.textBox3.Margin = new System.Windows.Forms.Padding(6);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.ShortcutsEnabled = false;
            this.textBox3.Size = new System.Drawing.Size(101, 22);
            this.textBox3.TabIndex = 120;
            this.textBox3.TabStop = false;
            this.textBox3.Text = "Size";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox18
            // 
            this.textBox18.BackColor = System.Drawing.SystemColors.Control;
            this.textBox18.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox18.Location = new System.Drawing.Point(22, 177);
            this.textBox18.Margin = new System.Windows.Forms.Padding(6);
            this.textBox18.Multiline = true;
            this.textBox18.Name = "textBox18";
            this.textBox18.ReadOnly = true;
            this.textBox18.ShortcutsEnabled = false;
            this.textBox18.Size = new System.Drawing.Size(178, 31);
            this.textBox18.TabIndex = 118;
            this.textBox18.TabStop = false;
            this.textBox18.Text = "Font";
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
            // dispFontSizeSheetNum
            // 
            this.dispFontSizeSheetNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispFontSizeSheetNum.Location = new System.Drawing.Point(211, 181);
            this.dispFontSizeSheetNum.Margin = new System.Windows.Forms.Padding(6);
            this.dispFontSizeSheetNum.MaxLength = 100;
            this.dispFontSizeSheetNum.Name = "dispFontSizeSheetNum";
            this.dispFontSizeSheetNum.Size = new System.Drawing.Size(132, 29);
            this.dispFontSizeSheetNum.TabIndex = 117;
            this.dispFontSizeSheetNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
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
            this.groupBox3.Controls.Add(this.setOutputFolder);
            this.groupBox3.Controls.Add(this.openOutputFolder);
            this.groupBox3.Controls.Add(this.dispOutputFolder);
            this.groupBox3.Controls.Add(this.editFilesSheetNum);
            this.groupBox3.Controls.Add(this.textBox19);
            this.groupBox3.Controls.Add(this.renameFilesCheck);
            this.groupBox3.Controls.Add(this.dispTotalDwgNum);
            this.groupBox3.Controls.Add(this.addSheetNumberCheck);
            this.groupBox3.Controls.Add(this.getFileInfo);
            this.groupBox3.Location = new System.Drawing.Point(15, 500);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox3.Size = new System.Drawing.Size(502, 432);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Add Sheet Number";
            // 
            // setOutputFolder
            // 
            this.setOutputFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setOutputFolder.Location = new System.Drawing.Point(15, 133);
            this.setOutputFolder.Margin = new System.Windows.Forms.Padding(6);
            this.setOutputFolder.Name = "setOutputFolder";
            this.setOutputFolder.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setOutputFolder.Size = new System.Drawing.Size(229, 46);
            this.setOutputFolder.TabIndex = 122;
            this.setOutputFolder.Text = "Set Output Folder";
            this.setOutputFolder.UseVisualStyleBackColor = true;
            // 
            // openOutputFolder
            // 
            this.openOutputFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.openOutputFolder.Location = new System.Drawing.Point(262, 133);
            this.openOutputFolder.Margin = new System.Windows.Forms.Padding(6);
            this.openOutputFolder.Name = "openOutputFolder";
            this.openOutputFolder.Size = new System.Drawing.Size(229, 46);
            this.openOutputFolder.TabIndex = 123;
            this.openOutputFolder.Text = "Open Output Folder";
            this.openOutputFolder.UseVisualStyleBackColor = true;
            // 
            // dispOutputFolder
            // 
            this.dispOutputFolder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOutputFolder.Location = new System.Drawing.Point(15, 190);
            this.dispOutputFolder.Margin = new System.Windows.Forms.Padding(6);
            this.dispOutputFolder.MaxLength = 1000;
            this.dispOutputFolder.Name = "dispOutputFolder";
            this.dispOutputFolder.Size = new System.Drawing.Size(475, 29);
            this.dispOutputFolder.TabIndex = 124;
            this.dispOutputFolder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // editFilesSheetNum
            // 
            this.editFilesSheetNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.editFilesSheetNum.Location = new System.Drawing.Point(14, 231);
            this.editFilesSheetNum.Margin = new System.Windows.Forms.Padding(6);
            this.editFilesSheetNum.Name = "editFilesSheetNum";
            this.editFilesSheetNum.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.editFilesSheetNum.Size = new System.Drawing.Size(477, 46);
            this.editFilesSheetNum.TabIndex = 112;
            this.editFilesSheetNum.Text = "Edit Files";
            this.editFilesSheetNum.UseVisualStyleBackColor = true;
            this.editFilesSheetNum.Click += new System.EventHandler(this.editFilesSheetNum_Click);
            // 
            // textBox19
            // 
            this.textBox19.BackColor = System.Drawing.SystemColors.Control;
            this.textBox19.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox19.Location = new System.Drawing.Point(22, 90);
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
            // renameFilesCheck
            // 
            this.renameFilesCheck.Checked = true;
            this.renameFilesCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.renameFilesCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.renameFilesCheck.Location = new System.Drawing.Point(22, 289);
            this.renameFilesCheck.Margin = new System.Windows.Forms.Padding(6);
            this.renameFilesCheck.Name = "renameFilesCheck";
            this.renameFilesCheck.Size = new System.Drawing.Size(240, 31);
            this.renameFilesCheck.TabIndex = 111;
            this.renameFilesCheck.Text = "Rename File";
            this.renameFilesCheck.UseVisualStyleBackColor = true;
            // 
            // dispTotalDwgNum
            // 
            this.dispTotalDwgNum.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispTotalDwgNum.Location = new System.Drawing.Point(211, 92);
            this.dispTotalDwgNum.Margin = new System.Windows.Forms.Padding(6);
            this.dispTotalDwgNum.MaxLength = 100;
            this.dispTotalDwgNum.Name = "dispTotalDwgNum";
            this.dispTotalDwgNum.Size = new System.Drawing.Size(132, 29);
            this.dispTotalDwgNum.TabIndex = 119;
            this.dispTotalDwgNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // addSheetNumberCheck
            // 
            this.addSheetNumberCheck.Checked = true;
            this.addSheetNumberCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.addSheetNumberCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.addSheetNumberCheck.Location = new System.Drawing.Point(22, 322);
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
            // DraftingPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ExcelTabControl);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "DraftingPane";
            this.Size = new System.Drawing.Size(550, 1532);
            this.ExcelTabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.testCoordinateGroup.ResumeLayout(false);
            this.testCoordinateGroup.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
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
        private System.Windows.Forms.GroupBox testCoordinateGroup;
        private System.Windows.Forms.Button testAddCoordinate;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox dispIncrement;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.ComboBox dispFontName;
        private System.Windows.Forms.Button setFontFolder;
        private System.Windows.Forms.TextBox dispFontPath;
        private System.Windows.Forms.TextBox dispValidCustomFont;
        private System.Windows.Forms.Button setOutputFolder;
        private System.Windows.Forms.Button openOutputFolder;
        private System.Windows.Forms.TextBox dispOutputFolder;
    }
}
