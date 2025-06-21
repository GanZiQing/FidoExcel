namespace ExcelAddIn2
{
    partial class DirectoryUserControl
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
            this.getDirectoryInfoGroup = new System.Windows.Forms.GroupBox();
            this.copyFiles = new System.Windows.Forms.Button();
            this.importSpecificFileNames = new System.Windows.Forms.Button();
            this.importFolderName = new System.Windows.Forms.Button();
            this.importFileName = new System.Windows.Forms.Button();
            this.addExtensionCheck = new System.Windows.Forms.CheckBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.renameFiles = new System.Windows.Forms.Button();
            this.importSpecificFile = new System.Windows.Forms.Button();
            this.dispExtension = new System.Windows.Forms.TextBox();
            this.labelExtension = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dirOpenPath = new System.Windows.Forms.Button();
            this.importFolderPath = new System.Windows.Forms.Button();
            this.nestedFoldersCheck = new System.Windows.Forms.CheckBox();
            this.setDirectory = new System.Windows.Forms.Button();
            this.dispDirectory = new System.Windows.Forms.TextBox();
            this.importFilePath = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.moveFiles = new System.Windows.Forms.Button();
            this.createFolders = new System.Windows.Forms.Button();
            this.mergeFoldersCheck = new System.Windows.Forms.CheckBox();
            this.appendFileNameCheck = new System.Windows.Forms.CheckBox();
            this.getDirectoryInfoGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // getDirectoryInfoGroup
            // 
            this.getDirectoryInfoGroup.Controls.Add(this.mergeFoldersCheck);
            this.getDirectoryInfoGroup.Controls.Add(this.appendFileNameCheck);
            this.getDirectoryInfoGroup.Controls.Add(this.createFolders);
            this.getDirectoryInfoGroup.Controls.Add(this.moveFiles);
            this.getDirectoryInfoGroup.Controls.Add(this.copyFiles);
            this.getDirectoryInfoGroup.Controls.Add(this.importSpecificFileNames);
            this.getDirectoryInfoGroup.Controls.Add(this.importFolderName);
            this.getDirectoryInfoGroup.Controls.Add(this.importFileName);
            this.getDirectoryInfoGroup.Controls.Add(this.addExtensionCheck);
            this.getDirectoryInfoGroup.Controls.Add(this.textBox7);
            this.getDirectoryInfoGroup.Controls.Add(this.renameFiles);
            this.getDirectoryInfoGroup.Controls.Add(this.importSpecificFile);
            this.getDirectoryInfoGroup.Controls.Add(this.dispExtension);
            this.getDirectoryInfoGroup.Controls.Add(this.labelExtension);
            this.getDirectoryInfoGroup.Controls.Add(this.textBox1);
            this.getDirectoryInfoGroup.Controls.Add(this.dirOpenPath);
            this.getDirectoryInfoGroup.Controls.Add(this.importFolderPath);
            this.getDirectoryInfoGroup.Controls.Add(this.nestedFoldersCheck);
            this.getDirectoryInfoGroup.Controls.Add(this.setDirectory);
            this.getDirectoryInfoGroup.Controls.Add(this.dispDirectory);
            this.getDirectoryInfoGroup.Controls.Add(this.importFilePath);
            this.getDirectoryInfoGroup.Location = new System.Drawing.Point(0, 0);
            this.getDirectoryInfoGroup.Margin = new System.Windows.Forms.Padding(6);
            this.getDirectoryInfoGroup.Name = "getDirectoryInfoGroup";
            this.getDirectoryInfoGroup.Padding = new System.Windows.Forms.Padding(6);
            this.getDirectoryInfoGroup.Size = new System.Drawing.Size(502, 653);
            this.getDirectoryInfoGroup.TabIndex = 5;
            this.getDirectoryInfoGroup.TabStop = false;
            this.getDirectoryInfoGroup.Text = "Get Directory Info";
            // 
            // copyFiles
            // 
            this.copyFiles.ForeColor = System.Drawing.SystemColors.WindowText;
            this.copyFiles.Location = new System.Drawing.Point(12, 518);
            this.copyFiles.Margin = new System.Windows.Forms.Padding(6);
            this.copyFiles.Name = "copyFiles";
            this.copyFiles.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.copyFiles.Size = new System.Drawing.Size(229, 46);
            this.copyFiles.TabIndex = 102;
            this.copyFiles.Text = "Copy Files/Folders";
            this.copyFiles.UseVisualStyleBackColor = true;
            this.copyFiles.Click += new System.EventHandler(this.copyFiles_Click);
            // 
            // importSpecificFileNames
            // 
            this.importSpecificFileNames.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importSpecificFileNames.Location = new System.Drawing.Point(261, 375);
            this.importSpecificFileNames.Margin = new System.Windows.Forms.Padding(6);
            this.importSpecificFileNames.Name = "importSpecificFileNames";
            this.importSpecificFileNames.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importSpecificFileNames.Size = new System.Drawing.Size(229, 46);
            this.importSpecificFileNames.TabIndex = 12;
            this.importSpecificFileNames.Text = "Get File Names w Type";
            this.importSpecificFileNames.UseVisualStyleBackColor = true;
            this.importSpecificFileNames.Click += new System.EventHandler(this.importSpecificFileNames_Click);
            // 
            // importFolderName
            // 
            this.importFolderName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importFolderName.Location = new System.Drawing.Point(260, 232);
            this.importFolderName.Margin = new System.Windows.Forms.Padding(6);
            this.importFolderName.Name = "importFolderName";
            this.importFolderName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importFolderName.Size = new System.Drawing.Size(229, 46);
            this.importFolderName.TabIndex = 9;
            this.importFolderName.Text = "Get Folder Names";
            this.importFolderName.UseVisualStyleBackColor = true;
            this.importFolderName.Click += new System.EventHandler(this.importFolderName_Click);
            // 
            // importFileName
            // 
            this.importFileName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importFileName.Location = new System.Drawing.Point(260, 175);
            this.importFileName.Margin = new System.Windows.Forms.Padding(6);
            this.importFileName.Name = "importFileName";
            this.importFileName.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importFileName.Size = new System.Drawing.Size(229, 46);
            this.importFileName.TabIndex = 8;
            this.importFileName.Text = "Get File Names";
            this.importFileName.UseVisualStyleBackColor = true;
            this.importFileName.Click += new System.EventHandler(this.importFileName_Click);
            // 
            // addExtensionCheck
            // 
            this.addExtensionCheck.Checked = true;
            this.addExtensionCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.addExtensionCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.addExtensionCheck.Location = new System.Drawing.Point(260, 133);
            this.addExtensionCheck.Margin = new System.Windows.Forms.Padding(6);
            this.addExtensionCheck.Name = "addExtensionCheck";
            this.addExtensionCheck.Size = new System.Drawing.Size(226, 30);
            this.addExtensionCheck.TabIndex = 5;
            this.addExtensionCheck.Text = "Include Extension";
            this.addExtensionCheck.UseVisualStyleBackColor = true;
            // 
            // textBox7
            // 
            this.textBox7.BackColor = System.Drawing.SystemColors.Control;
            this.textBox7.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox7.Location = new System.Drawing.Point(16, 432);
            this.textBox7.Margin = new System.Windows.Forms.Padding(6);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(275, 22);
            this.textBox7.TabIndex = 101;
            this.textBox7.TabStop = false;
            this.textBox7.Text = "Edit Directory";
            // 
            // renameFiles
            // 
            this.renameFiles.ForeColor = System.Drawing.SystemColors.WindowText;
            this.renameFiles.Location = new System.Drawing.Point(11, 460);
            this.renameFiles.Margin = new System.Windows.Forms.Padding(6);
            this.renameFiles.Name = "renameFiles";
            this.renameFiles.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.renameFiles.Size = new System.Drawing.Size(229, 46);
            this.renameFiles.TabIndex = 12;
            this.renameFiles.Text = "Rename Files/Folders";
            this.renameFiles.UseVisualStyleBackColor = true;
            this.renameFiles.Click += new System.EventHandler(this.renameFiles_Click);
            // 
            // importSpecificFile
            // 
            this.importSpecificFile.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importSpecificFile.Location = new System.Drawing.Point(11, 375);
            this.importSpecificFile.Margin = new System.Windows.Forms.Padding(6);
            this.importSpecificFile.Name = "importSpecificFile";
            this.importSpecificFile.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importSpecificFile.Size = new System.Drawing.Size(229, 46);
            this.importSpecificFile.TabIndex = 11;
            this.importSpecificFile.Text = "Get File Details w Type";
            this.importSpecificFile.UseVisualStyleBackColor = true;
            this.importSpecificFile.Click += new System.EventHandler(this.importSpecificFile_Click);
            // 
            // dispExtension
            // 
            this.dispExtension.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispExtension.Location = new System.Drawing.Point(205, 327);
            this.dispExtension.Margin = new System.Windows.Forms.Padding(6);
            this.dispExtension.MaxLength = 100;
            this.dispExtension.Name = "dispExtension";
            this.dispExtension.Size = new System.Drawing.Size(281, 29);
            this.dispExtension.TabIndex = 10;
            this.dispExtension.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelExtension
            // 
            this.labelExtension.BackColor = System.Drawing.SystemColors.Control;
            this.labelExtension.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelExtension.Location = new System.Drawing.Point(11, 332);
            this.labelExtension.Margin = new System.Windows.Forms.Padding(6);
            this.labelExtension.Name = "labelExtension";
            this.labelExtension.ReadOnly = true;
            this.labelExtension.Size = new System.Drawing.Size(183, 22);
            this.labelExtension.TabIndex = 38;
            this.labelExtension.TabStop = false;
            this.labelExtension.Text = "Specify Extension";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(11, 292);
            this.textBox1.Margin = new System.Windows.Forms.Padding(6);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(275, 22);
            this.textBox1.TabIndex = 36;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "Get Specific File Type:";
            // 
            // dirOpenPath
            // 
            this.dirOpenPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dirOpenPath.Location = new System.Drawing.Point(260, 35);
            this.dirOpenPath.Margin = new System.Windows.Forms.Padding(6);
            this.dirOpenPath.Name = "dirOpenPath";
            this.dirOpenPath.Size = new System.Drawing.Size(229, 46);
            this.dirOpenPath.TabIndex = 2;
            this.dirOpenPath.Text = "Open Folder";
            this.dirOpenPath.UseVisualStyleBackColor = true;
            // 
            // importFolderPath
            // 
            this.importFolderPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importFolderPath.Location = new System.Drawing.Point(11, 232);
            this.importFolderPath.Margin = new System.Windows.Forms.Padding(6);
            this.importFolderPath.Name = "importFolderPath";
            this.importFolderPath.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importFolderPath.Size = new System.Drawing.Size(229, 46);
            this.importFolderPath.TabIndex = 7;
            this.importFolderPath.Text = "Get Folder Details";
            this.importFolderPath.UseVisualStyleBackColor = true;
            this.importFolderPath.Click += new System.EventHandler(this.importFolderPath_Click);
            // 
            // nestedFoldersCheck
            // 
            this.nestedFoldersCheck.Checked = true;
            this.nestedFoldersCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.nestedFoldersCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.nestedFoldersCheck.Location = new System.Drawing.Point(16, 133);
            this.nestedFoldersCheck.Margin = new System.Windows.Forms.Padding(6);
            this.nestedFoldersCheck.Name = "nestedFoldersCheck";
            this.nestedFoldersCheck.Size = new System.Drawing.Size(240, 30);
            this.nestedFoldersCheck.TabIndex = 4;
            this.nestedFoldersCheck.Text = "Check nested folders";
            this.nestedFoldersCheck.UseVisualStyleBackColor = true;
            // 
            // setDirectory
            // 
            this.setDirectory.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setDirectory.Location = new System.Drawing.Point(11, 35);
            this.setDirectory.Margin = new System.Windows.Forms.Padding(6);
            this.setDirectory.Name = "setDirectory";
            this.setDirectory.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setDirectory.Size = new System.Drawing.Size(229, 46);
            this.setDirectory.TabIndex = 1;
            this.setDirectory.Text = "Set Folder";
            this.setDirectory.UseVisualStyleBackColor = true;
            // 
            // dispDirectory
            // 
            this.dispDirectory.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispDirectory.Location = new System.Drawing.Point(11, 92);
            this.dispDirectory.Margin = new System.Windows.Forms.Padding(6);
            this.dispDirectory.MaxLength = 1000;
            this.dispDirectory.Name = "dispDirectory";
            this.dispDirectory.Size = new System.Drawing.Size(475, 29);
            this.dispDirectory.TabIndex = 3;
            this.dispDirectory.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // importFilePath
            // 
            this.importFilePath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importFilePath.Location = new System.Drawing.Point(11, 174);
            this.importFilePath.Margin = new System.Windows.Forms.Padding(6);
            this.importFilePath.Name = "importFilePath";
            this.importFilePath.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importFilePath.Size = new System.Drawing.Size(229, 46);
            this.importFilePath.TabIndex = 6;
            this.importFilePath.Text = "Get File Details";
            this.importFilePath.UseVisualStyleBackColor = true;
            this.importFilePath.Click += new System.EventHandler(this.importFilePath_Click);
            // 
            // moveFiles
            // 
            this.moveFiles.ForeColor = System.Drawing.SystemColors.WindowText;
            this.moveFiles.Location = new System.Drawing.Point(261, 518);
            this.moveFiles.Margin = new System.Windows.Forms.Padding(6);
            this.moveFiles.Name = "moveFiles";
            this.moveFiles.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.moveFiles.Size = new System.Drawing.Size(229, 46);
            this.moveFiles.TabIndex = 104;
            this.moveFiles.Text = "Move Files/Folders";
            this.moveFiles.UseVisualStyleBackColor = true;
            this.moveFiles.Click += new System.EventHandler(this.moveFiles_Click);
            // 
            // createFolders
            // 
            this.createFolders.ForeColor = System.Drawing.SystemColors.WindowText;
            this.createFolders.Location = new System.Drawing.Point(261, 460);
            this.createFolders.Margin = new System.Windows.Forms.Padding(6);
            this.createFolders.Name = "createFolders";
            this.createFolders.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.createFolders.Size = new System.Drawing.Size(229, 46);
            this.createFolders.TabIndex = 105;
            this.createFolders.Text = "Create Folders";
            this.createFolders.UseVisualStyleBackColor = true;
            this.createFolders.Click += new System.EventHandler(this.createFolders_Click);
            // 
            // mergeFoldersCheck
            // 
            this.mergeFoldersCheck.Checked = true;
            this.mergeFoldersCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.mergeFoldersCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.mergeFoldersCheck.Location = new System.Drawing.Point(16, 618);
            this.mergeFoldersCheck.Margin = new System.Windows.Forms.Padding(6);
            this.mergeFoldersCheck.Name = "mergeFoldersCheck";
            this.mergeFoldersCheck.Size = new System.Drawing.Size(473, 30);
            this.mergeFoldersCheck.TabIndex = 107;
            this.mergeFoldersCheck.Text = "Merge folder if exist";
            this.mergeFoldersCheck.UseVisualStyleBackColor = true;
            // 
            // appendFileNameCheck
            // 
            this.appendFileNameCheck.Checked = true;
            this.appendFileNameCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.appendFileNameCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.appendFileNameCheck.Location = new System.Drawing.Point(16, 576);
            this.appendFileNameCheck.Margin = new System.Windows.Forms.Padding(6);
            this.appendFileNameCheck.Name = "appendFileNameCheck";
            this.appendFileNameCheck.Size = new System.Drawing.Size(470, 30);
            this.appendFileNameCheck.TabIndex = 106;
            this.appendFileNameCheck.Text = "Append (n) to file name if exist";
            this.appendFileNameCheck.UseVisualStyleBackColor = true;
            // 
            // DirectoryUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.getDirectoryInfoGroup);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "DirectoryUserControl";
            this.Size = new System.Drawing.Size(502, 657);
            this.getDirectoryInfoGroup.ResumeLayout(false);
            this.getDirectoryInfoGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox getDirectoryInfoGroup;
        private System.Windows.Forms.CheckBox addExtensionCheck;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.Button renameFiles;
        private System.Windows.Forms.Button importSpecificFile;
        private System.Windows.Forms.TextBox dispExtension;
        private System.Windows.Forms.TextBox labelExtension;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button dirOpenPath;
        private System.Windows.Forms.Button importFolderPath;
        private System.Windows.Forms.CheckBox nestedFoldersCheck;
        private System.Windows.Forms.Button setDirectory;
        private System.Windows.Forms.TextBox dispDirectory;
        private System.Windows.Forms.Button importFilePath;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button importFolderName;
        private System.Windows.Forms.Button importFileName;
        private System.Windows.Forms.Button importSpecificFileNames;
        private System.Windows.Forms.Button copyFiles;
        private System.Windows.Forms.Button moveFiles;
        private System.Windows.Forms.Button createFolders;
        private System.Windows.Forms.CheckBox mergeFoldersCheck;
        private System.Windows.Forms.CheckBox appendFileNameCheck;
    }
}
