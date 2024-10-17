namespace ScreenshotApp
{
    partial class SettingsForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dispFolderPath = new System.Windows.Forms.TextBox();
            this.openFolder = new System.Windows.Forms.Button();
            this.dispFilePath = new System.Windows.Forms.TextBox();
            this.overwriteCheck = new System.Windows.Forms.CheckBox();
            this.dispFileName = new System.Windows.Forms.TextBox();
            this.setFolder = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.clipboardCheck = new System.Windows.Forms.CheckBox();
            this.saveFileCheck = new System.Windows.Forms.CheckBox();
            this.alwaysOnTopCheck = new System.Windows.Forms.CheckBox();
            this.showHotKeyForm = new System.Windows.Forms.Button();
            this.aspectRatioCheck = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // dispFolderPath
            // 
            this.dispFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dispFolderPath.BackColor = System.Drawing.Color.White;
            this.dispFolderPath.Location = new System.Drawing.Point(153, 15);
            this.dispFolderPath.Name = "dispFolderPath";
            this.dispFolderPath.Size = new System.Drawing.Size(602, 20);
            this.dispFolderPath.TabIndex = 2;
            // 
            // openFolder
            // 
            this.openFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.openFolder.Location = new System.Drawing.Point(118, 114);
            this.openFolder.Name = "openFolder";
            this.openFolder.Size = new System.Drawing.Size(100, 25);
            this.openFolder.TabIndex = 5;
            this.openFolder.Text = "Open Folder";
            this.openFolder.UseVisualStyleBackColor = true;
            // 
            // dispFilePath
            // 
            this.dispFilePath.BackColor = System.Drawing.SystemColors.Control;
            this.dispFilePath.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dispFilePath.Location = new System.Drawing.Point(12, 43);
            this.dispFilePath.Multiline = true;
            this.dispFilePath.Name = "dispFilePath";
            this.dispFilePath.ReadOnly = true;
            this.dispFilePath.Size = new System.Drawing.Size(135, 17);
            this.dispFilePath.TabIndex = 9;
            this.dispFilePath.TabStop = false;
            this.dispFilePath.Text = "File Name:";
            this.dispFilePath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // overwriteCheck
            // 
            this.overwriteCheck.AutoSize = true;
            this.overwriteCheck.Location = new System.Drawing.Point(12, 66);
            this.overwriteCheck.Name = "overwriteCheck";
            this.overwriteCheck.Size = new System.Drawing.Size(134, 17);
            this.overwriteCheck.TabIndex = 4;
            this.overwriteCheck.Text = "Overwrite Existing Files";
            this.overwriteCheck.UseVisualStyleBackColor = true;
            // 
            // dispFileName
            // 
            this.dispFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dispFileName.BackColor = System.Drawing.Color.White;
            this.dispFileName.Location = new System.Drawing.Point(153, 40);
            this.dispFileName.Name = "dispFileName";
            this.dispFileName.Size = new System.Drawing.Size(602, 20);
            this.dispFileName.TabIndex = 3;
            this.dispFileName.Text = "Screenshot";
            // 
            // setFolder
            // 
            this.setFolder.Location = new System.Drawing.Point(12, 12);
            this.setFolder.Name = "setFolder";
            this.setFolder.Size = new System.Drawing.Size(135, 25);
            this.setFolder.TabIndex = 1;
            this.setFolder.Text = "Set Folder";
            this.setFolder.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Location = new System.Drawing.Point(655, 114);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 25);
            this.cancelButton.TabIndex = 7;
            this.cancelButton.Text = "Close";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // clipboardCheck
            // 
            this.clipboardCheck.AutoSize = true;
            this.clipboardCheck.Checked = true;
            this.clipboardCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.clipboardCheck.Location = new System.Drawing.Point(12, 89);
            this.clipboardCheck.Name = "clipboardCheck";
            this.clipboardCheck.Size = new System.Drawing.Size(109, 17);
            this.clipboardCheck.TabIndex = 10;
            this.clipboardCheck.Text = "Copy to Clipboard";
            this.clipboardCheck.UseVisualStyleBackColor = true;
            // 
            // saveFileCheck
            // 
            this.saveFileCheck.AutoSize = true;
            this.saveFileCheck.Checked = true;
            this.saveFileCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.saveFileCheck.Location = new System.Drawing.Point(127, 89);
            this.saveFileCheck.Name = "saveFileCheck";
            this.saveFileCheck.Size = new System.Drawing.Size(70, 17);
            this.saveFileCheck.TabIndex = 11;
            this.saveFileCheck.Text = "Save File";
            this.saveFileCheck.UseVisualStyleBackColor = true;
            // 
            // alwaysOnTopCheck
            // 
            this.alwaysOnTopCheck.AutoSize = true;
            this.alwaysOnTopCheck.Location = new System.Drawing.Point(203, 89);
            this.alwaysOnTopCheck.Name = "alwaysOnTopCheck";
            this.alwaysOnTopCheck.Size = new System.Drawing.Size(92, 17);
            this.alwaysOnTopCheck.TabIndex = 12;
            this.alwaysOnTopCheck.Text = "Always on top";
            this.alwaysOnTopCheck.UseVisualStyleBackColor = true;
            // 
            // showHotKeyForm
            // 
            this.showHotKeyForm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.showHotKeyForm.Location = new System.Drawing.Point(12, 114);
            this.showHotKeyForm.Name = "showHotKeyForm";
            this.showHotKeyForm.Size = new System.Drawing.Size(100, 25);
            this.showHotKeyForm.TabIndex = 13;
            this.showHotKeyForm.Text = "Show Hotkey";
            this.showHotKeyForm.UseVisualStyleBackColor = true;
            this.showHotKeyForm.Click += new System.EventHandler(this.ShowHotKeyForm_Click);
            // 
            // aspectRatioCheck
            // 
            this.aspectRatioCheck.AutoSize = true;
            this.aspectRatioCheck.Location = new System.Drawing.Point(301, 89);
            this.aspectRatioCheck.Name = "aspectRatioCheck";
            this.aspectRatioCheck.Size = new System.Drawing.Size(103, 17);
            this.aspectRatioCheck.TabIndex = 14;
            this.aspectRatioCheck.Text = "Fix Aspect Ratio";
            this.aspectRatioCheck.UseVisualStyleBackColor = true;
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 151);
            this.Controls.Add(this.aspectRatioCheck);
            this.Controls.Add(this.showHotKeyForm);
            this.Controls.Add(this.alwaysOnTopCheck);
            this.Controls.Add(this.saveFileCheck);
            this.Controls.Add(this.clipboardCheck);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.setFolder);
            this.Controls.Add(this.dispFileName);
            this.Controls.Add(this.overwriteCheck);
            this.Controls.Add(this.dispFilePath);
            this.Controls.Add(this.dispFolderPath);
            this.Controls.Add(this.openFolder);
            this.MaximumSize = new System.Drawing.Size(2000, 190);
            this.MinimumSize = new System.Drawing.Size(783, 190);
            this.Name = "SettingsForm";
            this.Text = "SettingsForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox dispFolderPath;
        private System.Windows.Forms.Button openFolder;
        private System.Windows.Forms.TextBox dispFilePath;
        private System.Windows.Forms.CheckBox overwriteCheck;
        private System.Windows.Forms.TextBox dispFileName;
        private System.Windows.Forms.Button setFolder;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.CheckBox clipboardCheck;
        private System.Windows.Forms.CheckBox saveFileCheck;
        private System.Windows.Forms.CheckBox alwaysOnTopCheck;
        private System.Windows.Forms.Button showHotKeyForm;
        private System.Windows.Forms.CheckBox aspectRatioCheck;
    }
}