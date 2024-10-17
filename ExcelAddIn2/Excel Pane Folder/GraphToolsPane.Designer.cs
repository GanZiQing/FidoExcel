namespace ExcelAddIn2.Excel_Pane_Folder
{
    partial class GraphToolsPane
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
            this.graphTabPage = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dispAmpFactor = new System.Windows.Forms.TextBox();
            this.terminateAtNullCheck2 = new System.Windows.Forms.CheckBox();
            this.dispDataSeries = new System.Windows.Forms.TextBox();
            this.setDataSeries = new System.Windows.Forms.Button();
            this.dispOutputRange = new System.Windows.Forms.TextBox();
            this.setOutputRange = new System.Windows.Forms.Button();
            this.setRange1 = new System.Windows.Forms.Button();
            this.runInterpolation = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.clearChartCheck = new System.Windows.Forms.CheckBox();
            this.terminateAtNullCheck = new System.Windows.Forms.CheckBox();
            this.clearChart = new System.Windows.Forms.Button();
            this.lineCheck = new System.Windows.Forms.CheckBox();
            this.pointCheck = new System.Windows.Forms.CheckBox();
            this.addSeries = new System.Windows.Forms.Button();
            this.setNameRange = new System.Windows.Forms.Button();
            this.dispNameRange = new System.Windows.Forms.TextBox();
            this.dispYRange = new System.Windows.Forms.TextBox();
            this.setYRange = new System.Windows.Forms.Button();
            this.dispXRange = new System.Windows.Forms.TextBox();
            this.setXRange = new System.Windows.Forms.Button();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.graphTabPage.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // graphTabPage
            // 
            this.graphTabPage.BackColor = System.Drawing.SystemColors.Control;
            this.graphTabPage.Controls.Add(this.groupBox1);
            this.graphTabPage.Controls.Add(this.groupBox2);
            this.graphTabPage.Location = new System.Drawing.Point(4, 22);
            this.graphTabPage.Name = "graphTabPage";
            this.graphTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.graphTabPage.Size = new System.Drawing.Size(289, 868);
            this.graphTabPage.TabIndex = 1;
            this.graphTabPage.Text = "Graph Tools";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dispAmpFactor);
            this.groupBox1.Controls.Add(this.terminateAtNullCheck2);
            this.groupBox1.Controls.Add(this.dispDataSeries);
            this.groupBox1.Controls.Add(this.setDataSeries);
            this.groupBox1.Controls.Add(this.dispOutputRange);
            this.groupBox1.Controls.Add(this.setOutputRange);
            this.groupBox1.Controls.Add(this.setRange1);
            this.groupBox1.Controls.Add(this.runInterpolation);
            this.groupBox1.Location = new System.Drawing.Point(6, 270);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(277, 199);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Interpolation (Linear)";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label2.Location = new System.Drawing.Point(6, 137);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(122, 20);
            this.label2.TabIndex = 38;
            this.label2.Text = "Amplification Factor";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dispAmpFactor
            // 
            this.dispAmpFactor.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispAmpFactor.Location = new System.Drawing.Point(134, 138);
            this.dispAmpFactor.Name = "dispAmpFactor";
            this.dispAmpFactor.Size = new System.Drawing.Size(134, 20);
            this.dispAmpFactor.TabIndex = 46;
            this.dispAmpFactor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispAmpFactor.WordWrap = false;
            // 
            // terminateAtNullCheck2
            // 
            this.terminateAtNullCheck2.AutoSize = true;
            this.terminateAtNullCheck2.Checked = true;
            this.terminateAtNullCheck2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.terminateAtNullCheck2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.terminateAtNullCheck2.Location = new System.Drawing.Point(6, 112);
            this.terminateAtNullCheck2.MinimumSize = new System.Drawing.Size(150, 17);
            this.terminateAtNullCheck2.Name = "terminateAtNullCheck2";
            this.terminateAtNullCheck2.Size = new System.Drawing.Size(150, 17);
            this.terminateAtNullCheck2.TabIndex = 44;
            this.terminateAtNullCheck2.Text = "Terminate ranges at null";
            this.terminateAtNullCheck2.UseVisualStyleBackColor = true;
            // 
            // dispDataSeries
            // 
            this.dispDataSeries.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispDataSeries.Location = new System.Drawing.Point(134, 84);
            this.dispDataSeries.Name = "dispDataSeries";
            this.dispDataSeries.Size = new System.Drawing.Size(134, 20);
            this.dispDataSeries.TabIndex = 41;
            this.dispDataSeries.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispDataSeries.WordWrap = false;
            // 
            // setDataSeries
            // 
            this.setDataSeries.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setDataSeries.Location = new System.Drawing.Point(6, 81);
            this.setDataSeries.Name = "setDataSeries";
            this.setDataSeries.Size = new System.Drawing.Size(122, 25);
            this.setDataSeries.TabIndex = 40;
            this.setDataSeries.Text = "Set Series Names";
            this.setDataSeries.UseVisualStyleBackColor = true;
            // 
            // dispOutputRange
            // 
            this.dispOutputRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispOutputRange.Location = new System.Drawing.Point(134, 53);
            this.dispOutputRange.Name = "dispOutputRange";
            this.dispOutputRange.Size = new System.Drawing.Size(134, 20);
            this.dispOutputRange.TabIndex = 43;
            this.dispOutputRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispOutputRange.WordWrap = false;
            // 
            // setOutputRange
            // 
            this.setOutputRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setOutputRange.Location = new System.Drawing.Point(6, 50);
            this.setOutputRange.Name = "setOutputRange";
            this.setOutputRange.Size = new System.Drawing.Size(122, 25);
            this.setOutputRange.TabIndex = 42;
            this.setOutputRange.Text = "Set Output Range";
            this.setOutputRange.UseVisualStyleBackColor = true;
            // 
            // setRange1
            // 
            this.setRange1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setRange1.Location = new System.Drawing.Point(6, 19);
            this.setRange1.Name = "setRange1";
            this.setRange1.Size = new System.Drawing.Size(122, 25);
            this.setRange1.TabIndex = 38;
            this.setRange1.Text = "Set Ranges";
            this.setRange1.UseVisualStyleBackColor = true;
            // 
            // runInterpolation
            // 
            this.runInterpolation.ForeColor = System.Drawing.SystemColors.WindowText;
            this.runInterpolation.Location = new System.Drawing.Point(6, 164);
            this.runInterpolation.Name = "runInterpolation";
            this.runInterpolation.Size = new System.Drawing.Size(262, 26);
            this.runInterpolation.TabIndex = 37;
            this.runInterpolation.Text = "Interpolate Ranges";
            this.runInterpolation.UseVisualStyleBackColor = true;
            this.runInterpolation.Click += new System.EventHandler(this.runInterpolation_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.clearChartCheck);
            this.groupBox2.Controls.Add(this.terminateAtNullCheck);
            this.groupBox2.Controls.Add(this.clearChart);
            this.groupBox2.Controls.Add(this.lineCheck);
            this.groupBox2.Controls.Add(this.pointCheck);
            this.groupBox2.Controls.Add(this.addSeries);
            this.groupBox2.Controls.Add(this.setNameRange);
            this.groupBox2.Controls.Add(this.dispNameRange);
            this.groupBox2.Controls.Add(this.dispYRange);
            this.groupBox2.Controls.Add(this.setYRange);
            this.groupBox2.Controls.Add(this.dispXRange);
            this.groupBox2.Controls.Add(this.setXRange);
            this.groupBox2.Location = new System.Drawing.Point(6, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(277, 258);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Add Chart";
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(3, 159);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(262, 20);
            this.label1.TabIndex = 44;
            this.label1.Text = "Set Formats:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // clearChartCheck
            // 
            this.clearChartCheck.AutoSize = true;
            this.clearChartCheck.Checked = true;
            this.clearChartCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.clearChartCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.clearChartCheck.Location = new System.Drawing.Point(0, 237);
            this.clearChartCheck.MinimumSize = new System.Drawing.Size(140, 17);
            this.clearChartCheck.Name = "clearChartCheck";
            this.clearChartCheck.Size = new System.Drawing.Size(140, 17);
            this.clearChartCheck.TabIndex = 43;
            this.clearChartCheck.Text = "Clear chart before plot";
            this.clearChartCheck.UseVisualStyleBackColor = true;
            // 
            // terminateAtNullCheck
            // 
            this.terminateAtNullCheck.AutoSize = true;
            this.terminateAtNullCheck.Checked = true;
            this.terminateAtNullCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.terminateAtNullCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.terminateAtNullCheck.Location = new System.Drawing.Point(6, 139);
            this.terminateAtNullCheck.MinimumSize = new System.Drawing.Size(150, 17);
            this.terminateAtNullCheck.Name = "terminateAtNullCheck";
            this.terminateAtNullCheck.Size = new System.Drawing.Size(150, 17);
            this.terminateAtNullCheck.TabIndex = 42;
            this.terminateAtNullCheck.Text = "Terminate ranges at null";
            this.terminateAtNullCheck.UseVisualStyleBackColor = true;
            // 
            // clearChart
            // 
            this.clearChart.ForeColor = System.Drawing.SystemColors.WindowText;
            this.clearChart.Location = new System.Drawing.Point(142, 78);
            this.clearChart.Name = "clearChart";
            this.clearChart.Size = new System.Drawing.Size(130, 35);
            this.clearChart.TabIndex = 41;
            this.clearChart.Text = "Clear Chart";
            this.clearChart.UseVisualStyleBackColor = true;
            this.clearChart.Click += new System.EventHandler(this.clearChart_Click);
            // 
            // lineCheck
            // 
            this.lineCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.lineCheck.Location = new System.Drawing.Point(100, 182);
            this.lineCheck.Name = "lineCheck";
            this.lineCheck.Size = new System.Drawing.Size(67, 17);
            this.lineCheck.TabIndex = 39;
            this.lineCheck.Text = "Plot line";
            this.lineCheck.UseVisualStyleBackColor = true;
            // 
            // pointCheck
            // 
            this.pointCheck.ForeColor = System.Drawing.SystemColors.WindowText;
            this.pointCheck.Location = new System.Drawing.Point(6, 182);
            this.pointCheck.Name = "pointCheck";
            this.pointCheck.Size = new System.Drawing.Size(76, 17);
            this.pointCheck.TabIndex = 38;
            this.pointCheck.Text = "Plot points";
            this.pointCheck.UseVisualStyleBackColor = true;
            // 
            // addSeries
            // 
            this.addSeries.ForeColor = System.Drawing.SystemColors.WindowText;
            this.addSeries.Location = new System.Drawing.Point(0, 205);
            this.addSeries.Name = "addSeries";
            this.addSeries.Size = new System.Drawing.Size(265, 26);
            this.addSeries.TabIndex = 37;
            this.addSeries.Text = "Add Series to Chart";
            this.addSeries.UseVisualStyleBackColor = true;
            this.addSeries.Click += new System.EventHandler(this.addSeries_Click);
            // 
            // setNameRange
            // 
            this.setNameRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setNameRange.Location = new System.Drawing.Point(6, 78);
            this.setNameRange.Name = "setNameRange";
            this.setNameRange.Size = new System.Drawing.Size(130, 35);
            this.setNameRange.TabIndex = 35;
            this.setNameRange.Text = "Set Name Column (Cell)";
            this.setNameRange.UseVisualStyleBackColor = true;
            // 
            // dispNameRange
            // 
            this.dispNameRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispNameRange.Location = new System.Drawing.Point(6, 113);
            this.dispNameRange.Name = "dispNameRange";
            this.dispNameRange.Size = new System.Drawing.Size(130, 20);
            this.dispNameRange.TabIndex = 34;
            this.dispNameRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispNameRange.WordWrap = false;
            // 
            // dispYRange
            // 
            this.dispYRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispYRange.Location = new System.Drawing.Point(142, 52);
            this.dispYRange.Name = "dispYRange";
            this.dispYRange.Size = new System.Drawing.Size(130, 20);
            this.dispYRange.TabIndex = 16;
            this.dispYRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispYRange.WordWrap = false;
            // 
            // setYRange
            // 
            this.setYRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setYRange.Location = new System.Drawing.Point(142, 19);
            this.setYRange.Name = "setYRange";
            this.setYRange.Size = new System.Drawing.Size(130, 35);
            this.setYRange.TabIndex = 15;
            this.setYRange.Text = "Set Y Column (Cell)";
            this.setYRange.UseVisualStyleBackColor = true;
            // 
            // dispXRange
            // 
            this.dispXRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.dispXRange.Location = new System.Drawing.Point(6, 52);
            this.dispXRange.Name = "dispXRange";
            this.dispXRange.Size = new System.Drawing.Size(130, 20);
            this.dispXRange.TabIndex = 14;
            this.dispXRange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.dispXRange.WordWrap = false;
            // 
            // setXRange
            // 
            this.setXRange.ForeColor = System.Drawing.SystemColors.WindowText;
            this.setXRange.Location = new System.Drawing.Point(6, 19);
            this.setXRange.Name = "setXRange";
            this.setXRange.Size = new System.Drawing.Size(130, 35);
            this.setXRange.TabIndex = 13;
            this.setXRange.Text = "Set X Range";
            this.setXRange.UseVisualStyleBackColor = true;
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.graphTabPage);
            this.tabControl.Location = new System.Drawing.Point(3, 3);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(297, 894);
            this.tabControl.TabIndex = 1;
            // 
            // GraphToolsPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl);
            this.Name = "GraphToolsPane";
            this.Size = new System.Drawing.Size(300, 900);
            this.graphTabPage.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabControl.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage graphTabPage;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button addSeries;
        private System.Windows.Forms.Button setNameRange;
        private System.Windows.Forms.TextBox dispNameRange;
        private System.Windows.Forms.TextBox dispYRange;
        private System.Windows.Forms.Button setYRange;
        private System.Windows.Forms.TextBox dispXRange;
        private System.Windows.Forms.Button setXRange;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button runInterpolation;
        private System.Windows.Forms.Button setRange1;
        private System.Windows.Forms.TextBox dispOutputRange;
        private System.Windows.Forms.Button setOutputRange;
        private System.Windows.Forms.CheckBox lineCheck;
        private System.Windows.Forms.CheckBox pointCheck;
        private System.Windows.Forms.Button clearChart;
        private System.Windows.Forms.TextBox dispDataSeries;
        private System.Windows.Forms.Button setDataSeries;
        private System.Windows.Forms.CheckBox terminateAtNullCheck;
        private System.Windows.Forms.CheckBox clearChartCheck;
        private System.Windows.Forms.CheckBox terminateAtNullCheck2;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox dispAmpFactor;
    }
}
