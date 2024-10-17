namespace ExcelAddIn2
{
    partial class CustomInputBox
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
            this.valueTextBox = new System.Windows.Forms.TextBox();
            this.instruction = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // valueTextBox
            // 
            this.valueTextBox.Location = new System.Drawing.Point(21, 46);
            this.valueTextBox.Name = "valueTextBox";
            this.valueTextBox.Size = new System.Drawing.Size(277, 20);
            this.valueTextBox.TabIndex = 0;
            // 
            // instruction
            // 
            this.instruction.AutoSize = true;
            this.instruction.Location = new System.Drawing.Point(18, 26);
            this.instruction.Name = "instruction";
            this.instruction.Size = new System.Drawing.Size(58, 13);
            this.instruction.TabIndex = 1;
            this.instruction.Text = "Edit Value:";
            // 
            // okButton
            // 
            this.okButton.Location = new System.Drawing.Point(112, 76);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(90, 26);
            this.okButton.TabIndex = 2;
            this.okButton.Text = "Ok";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(208, 76);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(90, 26);
            this.cancelButton.TabIndex = 3;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // InputBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(310, 114);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.instruction);
            this.Controls.Add(this.valueTextBox);
            this.Name = "InputBox";
            this.Text = "InputBox";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox valueTextBox;
        private System.Windows.Forms.Label instruction;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
    }
}