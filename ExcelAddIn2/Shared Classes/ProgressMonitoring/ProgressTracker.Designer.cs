namespace ExcelAddIn2
{
    partial class ProgressTracker
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
            this.ProgressBar1 = new System.Windows.Forms.ProgressBar();
            this.ProgressLabel = new System.Windows.Forms.Label();
            this.CancelButt = new System.Windows.Forms.Button();
            this.ThreadLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ProgressBar1
            // 
            this.ProgressBar1.Location = new System.Drawing.Point(12, 29);
            this.ProgressBar1.Name = "ProgressBar1";
            this.ProgressBar1.Size = new System.Drawing.Size(260, 23);
            this.ProgressBar1.TabIndex = 0;
            // 
            // ProgressLabel
            // 
            this.ProgressLabel.AutoSize = true;
            this.ProgressLabel.Location = new System.Drawing.Point(12, 13);
            this.ProgressLabel.Name = "ProgressLabel";
            this.ProgressLabel.Size = new System.Drawing.Size(21, 13);
            this.ProgressLabel.TabIndex = 1;
            this.ProgressLabel.Text = "0%";
            // 
            // CancelButt
            // 
            this.CancelButt.Location = new System.Drawing.Point(100, 88);
            this.CancelButt.Name = "CancelButt";
            this.CancelButt.Size = new System.Drawing.Size(100, 23);
            this.CancelButt.TabIndex = 2;
            this.CancelButt.Text = "Cancel";
            this.CancelButt.UseVisualStyleBackColor = true;
            this.CancelButt.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ThreadLabel
            // 
            this.ThreadLabel.Location = new System.Drawing.Point(12, 55);
            this.ThreadLabel.Name = "ThreadLabel";
            this.ThreadLabel.Size = new System.Drawing.Size(240, 30);
            this.ThreadLabel.TabIndex = 3;
            this.ThreadLabel.Text = "Executing...";
            // 
            // ProgressTracker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 121);
            this.Controls.Add(this.ThreadLabel);
            this.Controls.Add(this.CancelButt);
            this.Controls.Add(this.ProgressLabel);
            this.Controls.Add(this.ProgressBar1);
            this.Name = "ProgressTracker";
            this.Text = "ProgressBar";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar ProgressBar1;
        private System.Windows.Forms.Label ProgressLabel;
        private System.Windows.Forms.Button CancelButt;
        private System.Windows.Forms.Label ThreadLabel;
    }
}