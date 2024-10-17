using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    class CellFormatObject
    {
        #region Initialisation
        Label sampleCell;
        Button setFill;
        Button resetFill;
        Button setFont;
        Button resetFont;
        Button setBorder;
        Button resetBorder;
        Button getFromCell;
        Button applyToCell;

        public bool[] isDefault;// = new bool[] { true, true, true };
        public bool ignoreDefaults;
        public ColorDialog fillColorDialog = new ColorDialog();
        public ColorDialog fontColorDialog = new ColorDialog();
        public ColorDialog borderColorDialog = new ColorDialog();


        public CellFormatObject(Label sampleCell,Button setFill, Button resetFill, Button setFont, Button resetFont, 
            Button setBorder, Button resetBorder, Button getFromCell, Button applyToCell)
        {
            this.sampleCell = sampleCell;
            this.setFill = setFill;
            this.resetFill = resetFill;

            this.setFont = setFont;
            this.resetFont = resetFont;

            this.setBorder = setBorder;
            this.resetBorder = resetBorder;

            this.getFromCell = getFromCell;
            this.applyToCell = applyToCell;

            isDefault = new bool[] { true, true, true };
            ignoreDefaults = false;
            SubscribeToEvents();
        }

        private void SubscribeToEvents()
        {
            setFill.Click += new EventHandler(setFill_Click);
            resetFill.Click += new EventHandler(resetFill_Click);
            setFont.Click += new EventHandler(setFont_Click);
            resetFont.Click += new EventHandler(resetFont_Click);
            setBorder.Click += new EventHandler(setBorder_Click);
            resetBorder.Click += new EventHandler(resetBorder_Click);
            sampleCell.Paint += new PaintEventHandler(border_Paint);

            getFromCell.Click += new EventHandler(getFromCell_Click);
            applyToCell.Click += new EventHandler(applyToCell_Click);
        }
        private void border_Paint(object sender, PaintEventArgs e)
        {
            if (sampleCell.BorderStyle != BorderStyle.None)
            {
                ControlPaint.DrawBorder(e.Graphics, sampleCell.ClientRectangle, borderColorDialog.Color, ButtonBorderStyle.Solid);
            }
        }
        #endregion

        #region Set and Reset Colors
        private void setFill_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = fillColorDialog;
            int defaultIndex = 0;

            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                isDefault[defaultIndex] = false;
                sampleCell.BackColor = colorDialog.Color;
            }
        }
        private void resetFill_Click(object sender, EventArgs e)
        {
            int defaultIndex = 0;
            if (MessageBox.Show("Set fill colour to no fill?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                isDefault[defaultIndex] = true;
                sampleCell.BackColor = Control.DefaultBackColor;
            }
        }

        private void setFont_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = fontColorDialog;
            int defaultIndex = 1;

            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                isDefault[defaultIndex] = false;
                sampleCell.ForeColor = colorDialog.Color;
            }
        }
        private void resetFont_Click(object sender, EventArgs e)
        {
            int defaultIndex = 1;
            if (MessageBox.Show("Set font colour to default?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                isDefault[defaultIndex] = true;
                sampleCell.ForeColor = Control.DefaultForeColor;
            }
        }

        private void setBorder_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = borderColorDialog;
            int defaultIndex = 2;

            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                isDefault[defaultIndex] = false;
                sampleCell.BorderStyle = BorderStyle.FixedSingle;
                sampleCell.Invalidate();
            }
        }
        private void resetBorder_Click(object sender, EventArgs e)
        {
            int defaultIndex = 2;
            if (MessageBox.Show("Remove border to default?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                isDefault[defaultIndex] = true;
                sampleCell.BorderStyle = BorderStyle.None;
            }
        }
        #endregion

        private void getFromCell_Click(object sender, EventArgs e)
        {
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            if (activeCell == null)
            {
                return;
            }

            // Fill
            if (activeCell.Interior.ColorIndex == -4142)
            {
                isDefault[0] = true;
                sampleCell.BackColor = Control.DefaultBackColor;
            }
            else
            {
                isDefault[0] = false;
                (int R, int G, int B) = DecimalToRGB((double)activeCell.Interior.Color);
                fillColorDialog.Color = System.Drawing.Color.FromArgb(R, G, B);
                sampleCell.BackColor = fillColorDialog.Color;
            }

            // Font
            if (activeCell.Font.ColorIndex == -4105)
            {
                isDefault[1] = true;
                sampleCell.ForeColor = Control.DefaultForeColor;
            }
            else
            {
                isDefault[1] = false;
                (int R, int G, int B) = DecimalToRGB((double)activeCell.Font.Color);
                fontColorDialog.Color = System.Drawing.Color.FromArgb(R, G, B);
                sampleCell.ForeColor = fontColorDialog.Color;
            }

            // Border
            Border topBorder = activeCell.Borders[XlBordersIndex.xlEdgeTop];
            if (topBorder.ColorIndex == -4142)
            {
                isDefault[2] = true;
                sampleCell.BorderStyle = BorderStyle.None;
            }
            else
            {
                isDefault[2] = false;
                (int R, int G, int B) = DecimalToRGB((double)topBorder.Color);
                sampleCell.BorderStyle = BorderStyle.FixedSingle;
                borderColorDialog.Color = System.Drawing.Color.FromArgb(R, G, B);
                sampleCell.Invalidate();
            }
        }

        private (int, int, int) DecimalToRGB(double decimalColor)
        {
            //decimalColor = B * 65536 + G * 256 + R

            double B = Math.Floor(decimalColor/65536);
            double G = Math.Floor((decimalColor - B * 65536) / 256);
            double R = Math.Floor(decimalColor - B * 65536 - G * 256);

            return ((int)R, (int)G, (int)B);
        }

        private double RGBToDecimal(int R, int G, int B)
        {
            double decimalColor = B * 65536 + G * 256 + R;
            return decimalColor;
        }

        private void applyToCell_Click(object sender, EventArgs e)
        {
            Range selectedRange = Globals.ThisAddIn.Application.Selection;
            try
            {
                formatRange(selectedRange);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error formating selectd range\n\n + {ex}");
            }
        }

        public void formatRange(Range selectedRange)
        {
            // Fill
            if (isDefault[0])
            {
                if (!ignoreDefaults)
                {
                    selectedRange.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                }
            }
            else
            {
                selectedRange.Interior.Color = fillColorDialog.Color;
            }

            // Font
            if (isDefault[1])
            {
                if (!ignoreDefaults)
                {
                    selectedRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
                }
            }
            else
            {
                selectedRange.Font.Color = fontColorDialog.Color;
            }

            // Border
            if (isDefault[2])
            {
                if (!ignoreDefaults)
                {
                    selectedRange.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
                }
            }
            else
            {
                selectedRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                selectedRange.Borders.Color = borderColorDialog.Color;
            }
        }
    }
}
