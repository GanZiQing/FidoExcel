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
            this.addExtensionCheck = new System.Windows.Forms.CheckBox();
            this.insertRenameHeader = new System.Windows.Forms.Button();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.renameFiles = new System.Windows.Forms.Button();
            this.importSpecificFile = new System.Windows.Forms.Button();
            this.dispExtension = new System.Windows.Forms.TextBox();
            this.labelExtension = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.dirOpenPath = new System.Windows.Forms.Button();
            this.importFolderPath = new System.Windows.Forms.Button();
            this.checkNestedFolders = new System.Windows.Forms.CheckBox();
            this.setDirectory = new System.Windows.Forms.Button();
            this.dispDirectory = new System.Windows.Forms.TextBox();
            this.importFilePath = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.getDirectoryInfoGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // getDirectoryInfoGroup
            // 
            this.getDirectoryInfoGroup.Controls.Add(this.addExtensionCheck);
            this.getDirectoryInfoGroup.Controls.Add(this.insertRenameHeader);
            this.getDirectoryInfoGroup.Controls.Add(this.textBox7);
            this.getDirectoryInfoGroup.Controls.Add(this.renameFiles);
            this.getDirectoryInfoGroup.Controls.Add(this.importSpecificFile);
            this.getDirectoryInfoGroup.Controls.Add(this.dispExtension);
            this.getDirectoryInfoGroup.Controls.Add(this.labelExtension);
            this.getDirectoryInfoGroup.Controls.Add(this.textBox1);
            this.getDirectoryInfoGroup.Controls.Add(this.dirOpenPath);
            this.getDirectoryInfoGroup.Controls.Add(this.importFolderPath);
            this.getDirectoryInfoGroup.Controls.Add(this.checkNestedFolders);
            this.getDirectoryInfoGroup.Controls.Add(this.setDirectory);
            this.getDirectoryInfoGroup.Controls.Add(this.dispDirectory);
            this.getDirectoryInfoGroup.Controls.Add(this.importFilePath);
            this.getDirectoryInfoGroup.Location = new System.Drawing.Point(0, 0);
            this.getDirectoryInfoGroup.Margin = new System.Windows.Forms.Padding(6);
            this.getDirectoryInfoGroup.Name = "getDirectoryInfoGroup";
            this.getDirectoryInfoGroup.Padding = new System.Windows.Forms.Padding(6);
            this.getDirectoryInfoGroup.Size = new System.Drawing.Size(502, 462);
            this.getDirectoryInfoGroup.TabIndex = 5;
            this.getDirectoryInfoGroup.TabStop = false;
            this.getDirectoryInfoGroup.Text = "Get Directory Info";
            // 
            // addExtensionCheck
            // 
            this.addExtensionCheck.Checked = true;
            this.addExtensionCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.addExtensionCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.addExtensionCheck.Location = new System.Drawing.Point(260, 133);
            this.addExtensionCheck.Margin = new System.Windows.Forms.Padding(6);
            this.addExtensionCheck.Name = "addExtensionCheck";
            this.addExtensionCheck.Size = new System.Drawing.Size(226, 29);
            this.addExtensionCheck.TabIndex = 103;
            this.addExtensionCheck.Text = "Include Extension";
            this.addExtensionCheck.UseVisualStyleBackColor = true;
            // 
            // insertRenameHeader
            // 
            this.insertRenameHeader.ForeColor = System.Drawing.SystemColors.WindowText;
            this.insertRenameHeader.Location = new System.Drawing.Point(260, 399);
            this.insertRenameHeader.Margin = new System.Windows.Forms.Padding(6);
            this.insertRenameHeader.Name = "insertRenameHeader";
            this.insertRenameHeader.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.insertRenameHeader.Size = new System.Drawing.Size(229, 46);
            this.insertRenameHeader.TabIndex = 102;
            this.insertRenameHeader.Text = "Insert Header";
            this.insertRenameHeader.UseVisualStyleBackColor = true;
            // 
            // textBox7
            // 
            this.textBox7.BackColor = System.Drawing.SystemColors.Control;
            this.textBox7.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox7.Location = new System.Drawing.Point(16, 371);
            this.textBox7.Margin = new System.Windows.Forms.Padding(6);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(275, 22);
            this.textBox7.TabIndex = 101;
            this.textBox7.TabStop = false;
            this.textBox7.Text = "Rename Files";
            // 
            // renameFiles
            // 
            this.renameFiles.ForeColor = System.Drawing.SystemColors.WindowText;
            this.renameFiles.Location = new System.Drawing.Point(11, 399);
            this.renameFiles.Margin = new System.Windows.Forms.Padding(6);
            this.renameFiles.Name = "renameFiles";
            this.renameFiles.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.renameFiles.Size = new System.Drawing.Size(229, 46);
            this.renameFiles.TabIndex = 39;
            this.renameFiles.Text = "Rename Files";
            this.renameFiles.UseVisualStyleBackColor = true;
            // 
            // importSpecificFile
            // 
            this.importSpecificFile.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importSpecificFile.Location = new System.Drawing.Point(11, 314);
            this.importSpecificFile.Margin = new System.Windows.Forms.Padding(6);
            this.importSpecificFile.Name = "importSpecificFile";
            this.importSpecificFile.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importSpecificFile.Size = new System.Drawing.Size(478, 46);
            this.importSpecificFile.TabIndex = 16;
            this.importSpecificFile.Text = "Import Specific File Type";
            this.importSpecificFile.UseVisualStyleBackColor = true;
            this.importSpecificFile.Click += new System.EventHandler(this.importSpecificFile_Click);
            // 
            // dispExtension
            // 
            this.dispExtension.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispExtension.Location = new System.Drawing.Point(205, 266);
            this.dispExtension.Margin = new System.Windows.Forms.Padding(6);
            this.dispExtension.MaxLength = 100;
            this.dispExtension.Name = "dispExtension";
            this.dispExtension.Size = new System.Drawing.Size(281, 29);
            this.dispExtension.TabIndex = 15;
            this.dispExtension.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelExtension
            // 
            this.labelExtension.BackColor = System.Drawing.SystemColors.Control;
            this.labelExtension.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelExtension.Location = new System.Drawing.Point(11, 271);
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
            this.textBox1.Location = new System.Drawing.Point(11, 231);
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
            this.dirOpenPath.TabIndex = 10;
            this.dirOpenPath.Text = "Open Folder";
            this.dirOpenPath.UseVisualStyleBackColor = true;
            // 
            // importFolderPath
            // 
            this.importFolderPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importFolderPath.Location = new System.Drawing.Point(260, 174);
            this.importFolderPath.Margin = new System.Windows.Forms.Padding(6);
            this.importFolderPath.Name = "importFolderPath";
            this.importFolderPath.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.importFolderPath.Size = new System.Drawing.Size(229, 46);
            this.importFolderPath.TabIndex = 14;
            this.importFolderPath.Text = "Import All Folder Paths";
            this.importFolderPath.UseVisualStyleBackColor = true;
            this.importFolderPath.Click += new System.EventHandler(this.importFolderPath_Click);
            // 
            // checkNestedFolders
            // 
            this.checkNestedFolders.Checked = true;
            this.checkNestedFolders.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkNestedFolders.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkNestedFolders.Location = new System.Drawing.Point(16, 133);
            this.checkNestedFolders.Margin = new System.Windows.Forms.Padding(6);
            this.checkNestedFolders.Name = "checkNestedFolders";
            this.checkNestedFolders.Size = new System.Drawing.Size(240, 29);
            this.checkNestedFolders.TabIndex = 12;
            this.checkNestedFolders.Text = "Check nested folders";
            this.checkNestedFolders.UseVisualStyleBackColor = true;
            // 
            // setDirectory
            // 
            this.setDirectory.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setDirectory.Location = new System.Drawing.Point(11, 35);
            this.setDirectory.Margin = new System.Windows.Forms.Padding(6);
            this.setDirectory.Name = "setDirectory";
            this.setDirectory.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.setDirectory.Size = new System.Drawing.Size(229, 46);
            this.setDirectory.TabIndex = 9;
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
            this.dispDirectory.TabIndex = 11;
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
            this.importFilePath.TabIndex = 13;
            this.importFilePath.Text = "Import All File Paths";
            this.importFilePath.UseVisualStyleBackColor = true;
            this.importFilePath.Click += new System.EventHandler(this.importFilePath_Click);
            // 
            // DirectoryUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.getDirectoryInfoGroup);
            this.Name = "DirectoryUserControl";
            this.Size = new System.Drawing.Size(502, 462);
            this.getDirectoryInfoGroup.ResumeLayout(false);
            this.getDirectoryInfoGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox getDirectoryInfoGroup;
        private System.Windows.Forms.CheckBox addExtensionCheck;
        private System.Windows.Forms.Button insertRenameHeader;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.Button renameFiles;
        private System.Windows.Forms.Button importSpecificFile;
        private System.Windows.Forms.TextBox dispExtension;
        private System.Windows.Forms.TextBox labelExtension;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button dirOpenPath;
        private System.Windows.Forms.Button importFolderPath;
        private System.Windows.Forms.CheckBox checkNestedFolders;
        private System.Windows.Forms.Button setDirectory;
        private System.Windows.Forms.TextBox dispDirectory;
        private System.Windows.Forms.Button importFilePath;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}
