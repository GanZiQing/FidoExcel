namespace ScreenshotApp
{
    partial class HotKeyForm
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
            this.cancelButton = new System.Windows.Forms.Button();
            this.setHotKey = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dispKeyboardKey = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cancelButton.Location = new System.Drawing.Point(177, 83);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 25);
            this.cancelButton.TabIndex = 8;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // setHotKey
            // 
            this.setHotKey.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.setHotKey.Location = new System.Drawing.Point(71, 83);
            this.setHotKey.Name = "setHotKey";
            this.setHotKey.Size = new System.Drawing.Size(100, 25);
            this.setHotKey.TabIndex = 9;
            this.setHotKey.Text = "Set HotKey";
            this.setHotKey.UseVisualStyleBackColor = true;
            this.setHotKey.Click += new System.EventHandler(this.setHotKey_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(120, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Ctrl + Shift +";
            // 
            // dispKeyboardKey
            // 
            this.dispKeyboardKey.Location = new System.Drawing.Point(190, 40);
            this.dispKeyboardKey.Name = "dispKeyboardKey";
            this.dispKeyboardKey.Size = new System.Drawing.Size(23, 20);
            this.dispKeyboardKey.TabIndex = 11;
            this.dispKeyboardKey.Text = "A";
            this.dispKeyboardKey.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // HotKeyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(345, 120);
            this.Controls.Add(this.dispKeyboardKey);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.setHotKey);
            this.Controls.Add(this.cancelButton);
            this.Name = "HotKeyForm";
            this.Text = "HotKey";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button setHotKey;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox dispKeyboardKey;
    }
}