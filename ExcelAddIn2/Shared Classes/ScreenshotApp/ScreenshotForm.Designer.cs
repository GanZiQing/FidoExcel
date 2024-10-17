namespace ScreenshotApp
{
    partial class ScreenshotForm
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
            this.takeScreenshot = new System.Windows.Forms.Button();
            this.labelWidth = new System.Windows.Forms.TextBox();
            this.Settings = new System.Windows.Forms.Button();
            this.labelHeight = new System.Windows.Forms.TextBox();
            this.labelXPosition = new System.Windows.Forms.TextBox();
            this.lableYPosition = new System.Windows.Forms.TextBox();
            this.dispYPosition = new System.Windows.Forms.TextBox();
            this.dispXPosition = new System.Windows.Forms.TextBox();
            this.dispHeight = new System.Windows.Forms.TextBox();
            this.dispWidth = new System.Windows.Forms.TextBox();
            this.openFolder = new System.Windows.Forms.Button();
            this.openFile = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.dispStatus = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // takeScreenshot
            // 
            this.takeScreenshot.Location = new System.Drawing.Point(22, 22);
            this.takeScreenshot.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.takeScreenshot.Name = "takeScreenshot";
            this.takeScreenshot.Size = new System.Drawing.Size(183, 46);
            this.takeScreenshot.TabIndex = 0;
            this.takeScreenshot.Text = "Take Screenshot";
            this.takeScreenshot.UseVisualStyleBackColor = true;
            this.takeScreenshot.Click += new System.EventHandler(this.takeScreenshot_Click);
            // 
            // labelWidth
            // 
            this.labelWidth.BackColor = System.Drawing.SystemColors.Control;
            this.labelWidth.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelWidth.Location = new System.Drawing.Point(28, 80);
            this.labelWidth.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.labelWidth.Name = "labelWidth";
            this.labelWidth.ReadOnly = true;
            this.labelWidth.Size = new System.Drawing.Size(138, 22);
            this.labelWidth.TabIndex = 1;
            this.labelWidth.TabStop = false;
            this.labelWidth.Text = "Width:";
            // 
            // Settings
            // 
            this.Settings.Location = new System.Drawing.Point(605, 22);
            this.Settings.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Settings.Name = "Settings";
            this.Settings.Size = new System.Drawing.Size(183, 46);
            this.Settings.TabIndex = 1;
            this.Settings.Text = "Settings";
            this.Settings.UseVisualStyleBackColor = true;
            this.Settings.Click += new System.EventHandler(this.Settings_Click);
            // 
            // labelHeight
            // 
            this.labelHeight.BackColor = System.Drawing.SystemColors.Control;
            this.labelHeight.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelHeight.Location = new System.Drawing.Point(28, 111);
            this.labelHeight.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.labelHeight.Name = "labelHeight";
            this.labelHeight.ReadOnly = true;
            this.labelHeight.Size = new System.Drawing.Size(138, 22);
            this.labelHeight.TabIndex = 4;
            this.labelHeight.TabStop = false;
            this.labelHeight.Text = "Height:";
            // 
            // labelXPosition
            // 
            this.labelXPosition.BackColor = System.Drawing.SystemColors.Control;
            this.labelXPosition.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.labelXPosition.Location = new System.Drawing.Point(28, 142);
            this.labelXPosition.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.labelXPosition.Name = "labelXPosition";
            this.labelXPosition.ReadOnly = true;
            this.labelXPosition.Size = new System.Drawing.Size(138, 22);
            this.labelXPosition.TabIndex = 5;
            this.labelXPosition.TabStop = false;
            this.labelXPosition.Text = "X Position:";
            // 
            // lableYPosition
            // 
            this.lableYPosition.BackColor = System.Drawing.SystemColors.Control;
            this.lableYPosition.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lableYPosition.Location = new System.Drawing.Point(28, 174);
            this.lableYPosition.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.lableYPosition.Name = "lableYPosition";
            this.lableYPosition.ReadOnly = true;
            this.lableYPosition.Size = new System.Drawing.Size(138, 22);
            this.lableYPosition.TabIndex = 6;
            this.lableYPosition.TabStop = false;
            this.lableYPosition.Text = "Y Position:";
            // 
            // dispYPosition
            // 
            this.dispYPosition.BackColor = System.Drawing.Color.White;
            this.dispYPosition.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dispYPosition.Location = new System.Drawing.Point(176, 174);
            this.dispYPosition.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.dispYPosition.Name = "dispYPosition";
            this.dispYPosition.Size = new System.Drawing.Size(138, 22);
            this.dispYPosition.TabIndex = 5;
            // 
            // dispXPosition
            // 
            this.dispXPosition.BackColor = System.Drawing.Color.White;
            this.dispXPosition.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dispXPosition.Location = new System.Drawing.Point(176, 142);
            this.dispXPosition.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.dispXPosition.Name = "dispXPosition";
            this.dispXPosition.Size = new System.Drawing.Size(138, 22);
            this.dispXPosition.TabIndex = 4;
            // 
            // dispHeight
            // 
            this.dispHeight.BackColor = System.Drawing.Color.White;
            this.dispHeight.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dispHeight.Location = new System.Drawing.Point(176, 111);
            this.dispHeight.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.dispHeight.Name = "dispHeight";
            this.dispHeight.Size = new System.Drawing.Size(138, 22);
            this.dispHeight.TabIndex = 3;
            // 
            // dispWidth
            // 
            this.dispWidth.BackColor = System.Drawing.Color.White;
            this.dispWidth.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dispWidth.Location = new System.Drawing.Point(176, 80);
            this.dispWidth.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.dispWidth.Name = "dispWidth";
            this.dispWidth.Size = new System.Drawing.Size(138, 22);
            this.dispWidth.TabIndex = 2;
            // 
            // openFolder
            // 
            this.openFolder.Location = new System.Drawing.Point(216, 22);
            this.openFolder.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.openFolder.Name = "openFolder";
            this.openFolder.Size = new System.Drawing.Size(183, 46);
            this.openFolder.TabIndex = 8;
            this.openFolder.Text = "Open Folder";
            this.openFolder.UseVisualStyleBackColor = true;
            this.openFolder.Click += new System.EventHandler(this.openFolder_Click);
            // 
            // openFile
            // 
            this.openFile.Location = new System.Drawing.Point(411, 22);
            this.openFile.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.openFile.Name = "openFile";
            this.openFile.Size = new System.Drawing.Size(183, 46);
            this.openFile.TabIndex = 9;
            this.openFile.Text = "Open File";
            this.openFile.UseVisualStyleBackColor = true;
            this.openFile.Click += new System.EventHandler(this.openFile_Click);
            // 
            // closeButton
            // 
            this.closeButton.Location = new System.Drawing.Point(22, 208);
            this.closeButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(183, 46);
            this.closeButton.TabIndex = 11;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // dispStatus
            // 
            this.dispStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dispStatus.Location = new System.Drawing.Point(324, 80);
            this.dispStatus.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.dispStatus.Name = "dispStatus";
            this.dispStatus.Size = new System.Drawing.Size(1121, 678);
            this.dispStatus.TabIndex = 12;
            this.dispStatus.Text = "Status";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(800, 22);
            this.button1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(183, 46);
            this.button1.TabIndex = 14;
            this.button1.Text = "Device Info";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.showScreenInfo_Click);
            // 
            // ScreenshotForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1467, 774);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dispStatus);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.openFile);
            this.Controls.Add(this.openFolder);
            this.Controls.Add(this.dispYPosition);
            this.Controls.Add(this.dispXPosition);
            this.Controls.Add(this.dispHeight);
            this.Controls.Add(this.dispWidth);
            this.Controls.Add(this.lableYPosition);
            this.Controls.Add(this.labelXPosition);
            this.Controls.Add(this.labelHeight);
            this.Controls.Add(this.Settings);
            this.Controls.Add(this.labelWidth);
            this.Controls.Add(this.takeScreenshot);
            this.DoubleBuffered = true;
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "ScreenshotForm";
            this.Opacity = 0.8D;
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Screenshot Tool";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button takeScreenshot;
        private System.Windows.Forms.TextBox labelWidth;
        private System.Windows.Forms.Button Settings;
        private System.Windows.Forms.TextBox labelHeight;
        private System.Windows.Forms.TextBox labelXPosition;
        private System.Windows.Forms.TextBox lableYPosition;
        private System.Windows.Forms.TextBox dispYPosition;
        private System.Windows.Forms.TextBox dispXPosition;
        private System.Windows.Forms.TextBox dispHeight;
        private System.Windows.Forms.TextBox dispWidth;
        private System.Windows.Forms.Button openFolder;
        private System.Windows.Forms.Button openFile;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Label dispStatus;
        private System.Windows.Forms.Button button1;
    }


}

