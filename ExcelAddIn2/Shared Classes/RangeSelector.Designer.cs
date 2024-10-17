namespace ExcelAddIn2
{
    partial class RangeSelector
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
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.rangeListBox = new System.Windows.Forms.ListBox();
            this.deleteButton = new System.Windows.Forms.Button();
            this.LeftLabel = new System.Windows.Forms.Label();
            this.addRange = new System.Windows.Forms.Button();
            this.clearButton = new System.Windows.Forms.Button();
            this.moveToBottom = new System.Windows.Forms.Button();
            this.moveDown = new System.Windows.Forms.Button();
            this.moveUp = new System.Windows.Forms.Button();
            this.moveToTop = new System.Windows.Forms.Button();
            this.editButton = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.offSet = new System.Windows.Forms.Button();
            this.copy = new System.Windows.Forms.Button();
            this.useSameOffsetCheck = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // okButton
            // 
            this.okButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.okButton.Location = new System.Drawing.Point(64, 413);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(90, 25);
            this.okButton.TabIndex = 18;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.cancelButton.Location = new System.Drawing.Point(157, 413);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(90, 25);
            this.cancelButton.TabIndex = 17;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // rangeListBox
            // 
            this.rangeListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rangeListBox.FormattingEnabled = true;
            this.rangeListBox.Location = new System.Drawing.Point(13, 25);
            this.rangeListBox.Name = "rangeListBox";
            this.rangeListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.rangeListBox.Size = new System.Drawing.Size(241, 290);
            this.rangeListBox.TabIndex = 16;
            // 
            // deleteButton
            // 
            this.deleteButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.deleteButton.Location = new System.Drawing.Point(109, 328);
            this.deleteButton.Name = "deleteButton";
            this.deleteButton.Size = new System.Drawing.Size(90, 25);
            this.deleteButton.TabIndex = 24;
            this.deleteButton.Text = "Delete";
            this.deleteButton.UseVisualStyleBackColor = true;
            this.deleteButton.Click += new System.EventHandler(this.deleteButton_Click);
            // 
            // LeftLabel
            // 
            this.LeftLabel.AutoSize = true;
            this.LeftLabel.Location = new System.Drawing.Point(12, 9);
            this.LeftLabel.Name = "LeftLabel";
            this.LeftLabel.Size = new System.Drawing.Size(89, 13);
            this.LeftLabel.TabIndex = 14;
            this.LeftLabel.Text = "Selected Ranges";
            // 
            // addRange
            // 
            this.addRange.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.addRange.Location = new System.Drawing.Point(13, 328);
            this.addRange.Name = "addRange";
            this.addRange.Size = new System.Drawing.Size(90, 25);
            this.addRange.TabIndex = 25;
            this.addRange.Text = "Add Range";
            this.addRange.UseVisualStyleBackColor = true;
            this.addRange.Click += new System.EventHandler(this.addRange_Click);
            // 
            // clearButton
            // 
            this.clearButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.clearButton.Location = new System.Drawing.Point(205, 328);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(90, 25);
            this.clearButton.TabIndex = 26;
            this.clearButton.Text = "Clear All";
            this.clearButton.UseVisualStyleBackColor = true;
            this.clearButton.Click += new System.EventHandler(this.clearButton_Click);
            // 
            // moveToBottom
            // 
            this.moveToBottom.Location = new System.Drawing.Point(0, 110);
            this.moveToBottom.Margin = new System.Windows.Forms.Padding(0);
            this.moveToBottom.Name = "moveToBottom";
            this.moveToBottom.Size = new System.Drawing.Size(35, 25);
            this.moveToBottom.TabIndex = 30;
            this.moveToBottom.Text = "▼▼";
            this.moveToBottom.UseVisualStyleBackColor = true;
            this.moveToBottom.Click += new System.EventHandler(this.moveToBottom_Click);
            // 
            // moveDown
            // 
            this.moveDown.Location = new System.Drawing.Point(0, 80);
            this.moveDown.Name = "moveDown";
            this.moveDown.Size = new System.Drawing.Size(35, 25);
            this.moveDown.TabIndex = 29;
            this.moveDown.Text = "▼";
            this.moveDown.UseVisualStyleBackColor = true;
            this.moveDown.Click += new System.EventHandler(this.moveDown_Click);
            // 
            // moveUp
            // 
            this.moveUp.Location = new System.Drawing.Point(0, 30);
            this.moveUp.Name = "moveUp";
            this.moveUp.Size = new System.Drawing.Size(35, 25);
            this.moveUp.TabIndex = 28;
            this.moveUp.Text = "▲";
            this.moveUp.UseVisualStyleBackColor = true;
            this.moveUp.Click += new System.EventHandler(this.moveUp_Click);
            // 
            // moveToTop
            // 
            this.moveToTop.Location = new System.Drawing.Point(0, 0);
            this.moveToTop.Name = "moveToTop";
            this.moveToTop.Size = new System.Drawing.Size(35, 25);
            this.moveToTop.TabIndex = 27;
            this.moveToTop.Text = "▲▲";
            this.moveToTop.UseVisualStyleBackColor = true;
            this.moveToTop.Click += new System.EventHandler(this.moveToTop_Click);
            // 
            // editButton
            // 
            this.editButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.editButton.Location = new System.Drawing.Point(13, 359);
            this.editButton.Name = "editButton";
            this.editButton.Size = new System.Drawing.Size(90, 25);
            this.editButton.TabIndex = 31;
            this.editButton.Text = "Edit";
            this.editButton.UseVisualStyleBackColor = true;
            this.editButton.Click += new System.EventHandler(this.editButton_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.panel1.Controls.Add(this.moveToTop);
            this.panel1.Controls.Add(this.moveUp);
            this.panel1.Controls.Add(this.moveToBottom);
            this.panel1.Controls.Add(this.moveDown);
            this.panel1.Location = new System.Drawing.Point(260, 110);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(35, 135);
            this.panel1.TabIndex = 32;
            // 
            // offSet
            // 
            this.offSet.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.offSet.Location = new System.Drawing.Point(109, 359);
            this.offSet.Name = "offSet";
            this.offSet.Size = new System.Drawing.Size(90, 25);
            this.offSet.TabIndex = 33;
            this.offSet.Text = "Offset";
            this.offSet.UseVisualStyleBackColor = true;
            this.offSet.Click += new System.EventHandler(this.offSet_Click);
            // 
            // copy
            // 
            this.copy.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.copy.Location = new System.Drawing.Point(205, 359);
            this.copy.Name = "copy";
            this.copy.Size = new System.Drawing.Size(90, 25);
            this.copy.TabIndex = 34;
            this.copy.Text = "Copy";
            this.copy.UseVisualStyleBackColor = true;
            this.copy.Click += new System.EventHandler(this.copy_Click);
            // 
            // useSameOffsetCheck
            // 
            this.useSameOffsetCheck.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.useSameOffsetCheck.AutoSize = true;
            this.useSameOffsetCheck.Location = new System.Drawing.Point(109, 390);
            this.useSameOffsetCheck.Name = "useSameOffsetCheck";
            this.useSameOffsetCheck.Size = new System.Drawing.Size(109, 17);
            this.useSameOffsetCheck.TabIndex = 35;
            this.useSameOffsetCheck.Text = "Use Global Offset";
            this.useSameOffsetCheck.UseVisualStyleBackColor = true;
            // 
            // RangeSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(307, 450);
            this.Controls.Add(this.useSameOffsetCheck);
            this.Controls.Add(this.copy);
            this.Controls.Add(this.offSet);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.editButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.addRange);
            this.Controls.Add(this.deleteButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.rangeListBox);
            this.Controls.Add(this.LeftLabel);
            this.Name = "RangeSelector";
            this.Text = "Select Ranges";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.ListBox rangeListBox;
        private System.Windows.Forms.Button deleteButton;
        private System.Windows.Forms.Label LeftLabel;
        private System.Windows.Forms.Button addRange;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button moveToBottom;
        private System.Windows.Forms.Button moveDown;
        private System.Windows.Forms.Button moveUp;
        private System.Windows.Forms.Button moveToTop;
        private System.Windows.Forms.Button editButton;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button offSet;
        private System.Windows.Forms.Button copy;
        private System.Windows.Forms.CheckBox useSameOffsetCheck;
    }
}