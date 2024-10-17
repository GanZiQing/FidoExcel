namespace ExcelAddIn2
{
    partial class SheetSelector
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
            this.label1 = new System.Windows.Forms.Label();
            this.LeftLabel = new System.Windows.Forms.Label();
            this.RightLabel = new System.Windows.Forms.Label();
            this.LeftListBox = new System.Windows.Forms.ListBox();
            this.MyCancelButton = new System.Windows.Forms.Button();
            this.ConfirmationButton = new System.Windows.Forms.Button();
            this.RightListBox = new System.Windows.Forms.ListBox();
            this.MoveAllRight = new System.Windows.Forms.Button();
            this.MoveSelectionRight = new System.Windows.Forms.Button();
            this.MoveSelectionLeft = new System.Windows.Forms.Button();
            this.MoveAllLeft = new System.Windows.Forms.Button();
            this.DeleteButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Sheets";
            // 
            // LeftLabel
            // 
            this.LeftLabel.AutoSize = true;
            this.LeftLabel.Location = new System.Drawing.Point(12, 36);
            this.LeftLabel.Name = "LeftLabel";
            this.LeftLabel.Size = new System.Drawing.Size(93, 13);
            this.LeftLabel.TabIndex = 1;
            this.LeftLabel.Text = "Remaining Sheets";
            // 
            // RightLabel
            // 
            this.RightLabel.AutoSize = true;
            this.RightLabel.Location = new System.Drawing.Point(374, 36);
            this.RightLabel.Name = "RightLabel";
            this.RightLabel.Size = new System.Drawing.Size(85, 13);
            this.RightLabel.TabIndex = 2;
            this.RightLabel.Text = "Selected Sheets";
            // 
            // LeftListBox
            // 
            this.LeftListBox.FormattingEnabled = true;
            this.LeftListBox.Location = new System.Drawing.Point(15, 52);
            this.LeftListBox.Name = "LeftListBox";
            this.LeftListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.LeftListBox.Size = new System.Drawing.Size(320, 303);
            this.LeftListBox.TabIndex = 3;
            // 
            // MyCancelButton
            // 
            this.MyCancelButton.Location = new System.Drawing.Point(607, 362);
            this.MyCancelButton.Name = "MyCancelButton";
            this.MyCancelButton.Size = new System.Drawing.Size(90, 25);
            this.MyCancelButton.TabIndex = 5;
            this.MyCancelButton.Text = "Cancel";
            this.MyCancelButton.UseVisualStyleBackColor = true;
            this.MyCancelButton.Click += new System.EventHandler(this.MyCancelButton_Click);
            // 
            // ConfirmationButton
            // 
            this.ConfirmationButton.Location = new System.Drawing.Point(511, 362);
            this.ConfirmationButton.Name = "ConfirmationButton";
            this.ConfirmationButton.Size = new System.Drawing.Size(90, 25);
            this.ConfirmationButton.TabIndex = 6;
            this.ConfirmationButton.Text = "OK";
            this.ConfirmationButton.UseVisualStyleBackColor = true;
            this.ConfirmationButton.Click += new System.EventHandler(this.ConfirmationButton_Click);
            // 
            // RightListBox
            // 
            this.RightListBox.FormattingEnabled = true;
            this.RightListBox.Location = new System.Drawing.Point(377, 52);
            this.RightListBox.Name = "RightListBox";
            this.RightListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.RightListBox.Size = new System.Drawing.Size(320, 303);
            this.RightListBox.TabIndex = 7;
            // 
            // MoveAllRight
            // 
            this.MoveAllRight.Location = new System.Drawing.Point(341, 107);
            this.MoveAllRight.Name = "MoveAllRight";
            this.MoveAllRight.Size = new System.Drawing.Size(30, 25);
            this.MoveAllRight.TabIndex = 8;
            this.MoveAllRight.Text = ">>";
            this.MoveAllRight.UseVisualStyleBackColor = true;
            this.MoveAllRight.Click += new System.EventHandler(this.MoveAllRight_Click);
            // 
            // MoveSelectionRight
            // 
            this.MoveSelectionRight.Location = new System.Drawing.Point(341, 138);
            this.MoveSelectionRight.Name = "MoveSelectionRight";
            this.MoveSelectionRight.Size = new System.Drawing.Size(30, 25);
            this.MoveSelectionRight.TabIndex = 9;
            this.MoveSelectionRight.Text = ">";
            this.MoveSelectionRight.UseVisualStyleBackColor = true;
            this.MoveSelectionRight.Click += new System.EventHandler(this.MoveSelectionRight_Click);
            // 
            // MoveSelectionLeft
            // 
            this.MoveSelectionLeft.Location = new System.Drawing.Point(341, 210);
            this.MoveSelectionLeft.Name = "MoveSelectionLeft";
            this.MoveSelectionLeft.Size = new System.Drawing.Size(30, 25);
            this.MoveSelectionLeft.TabIndex = 10;
            this.MoveSelectionLeft.Text = "<";
            this.MoveSelectionLeft.UseVisualStyleBackColor = true;
            this.MoveSelectionLeft.Click += new System.EventHandler(this.MoveSelectionLeft_Click);
            // 
            // MoveAllLeft
            // 
            this.MoveAllLeft.Location = new System.Drawing.Point(341, 241);
            this.MoveAllLeft.Name = "MoveAllLeft";
            this.MoveAllLeft.Size = new System.Drawing.Size(30, 25);
            this.MoveAllLeft.TabIndex = 11;
            this.MoveAllLeft.Text = "<<";
            this.MoveAllLeft.UseVisualStyleBackColor = true;
            this.MoveAllLeft.Click += new System.EventHandler(this.MoveAllLeft_Click);
            // 
            // DeleteButton
            // 
            this.DeleteButton.Location = new System.Drawing.Point(415, 362);
            this.DeleteButton.Name = "DeleteButton";
            this.DeleteButton.Size = new System.Drawing.Size(90, 25);
            this.DeleteButton.TabIndex = 12;
            this.DeleteButton.Text = "Delete";
            this.DeleteButton.UseVisualStyleBackColor = true;
            this.DeleteButton.Visible = false;
            this.DeleteButton.Click += new System.EventHandler(this.DeleteButton_Click);
            // 
            // SheetSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 399);
            this.Controls.Add(this.DeleteButton);
            this.Controls.Add(this.MoveAllLeft);
            this.Controls.Add(this.MoveSelectionLeft);
            this.Controls.Add(this.MoveSelectionRight);
            this.Controls.Add(this.MoveAllRight);
            this.Controls.Add(this.RightListBox);
            this.Controls.Add(this.ConfirmationButton);
            this.Controls.Add(this.MyCancelButton);
            this.Controls.Add(this.LeftListBox);
            this.Controls.Add(this.RightLabel);
            this.Controls.Add(this.LeftLabel);
            this.Controls.Add(this.label1);
            this.Name = "SheetSelector";
            this.Text = "SheetSelector";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label LeftLabel;
        private System.Windows.Forms.Label RightLabel;
        private System.Windows.Forms.ListBox LeftListBox;
        private System.Windows.Forms.Button MyCancelButton;
        private System.Windows.Forms.Button ConfirmationButton;
        private System.Windows.Forms.ListBox RightListBox;
        private System.Windows.Forms.Button MoveAllRight;
        private System.Windows.Forms.Button MoveSelectionRight;
        private System.Windows.Forms.Button MoveSelectionLeft;
        private System.Windows.Forms.Button MoveAllLeft;
        private System.Windows.Forms.Button DeleteButton;
    }
}