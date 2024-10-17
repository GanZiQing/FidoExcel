using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Excel.Application;
using static ExcelAddIn2.CommonUtilities;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using PdfSharp.Snippets.Drawing;
using System.Windows.Media.Media3D;
using XlColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class FormatToolsPane : UserControl
    {
        #region Init
        Application thisApp = Globals.ThisAddIn.Application;
        Dictionary<string, AttributeTextBox> TextAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();

        public FormatToolsPane()
        {
            InitializeComponent();
            CreateAttributes();
            CreateCellFormatObjects();
            AddToolTips();
        }

        private void CreateAttributes()
        {
            // Create Attribute Objects 
            #region Format Table
            RangeTextBox CompCol = new RangeTextBox("CompCol", DispCompCol, SetCompCol, "range");
            TextAttributeDic.Add("CompCol", CompCol);
            #endregion

            #region Compare Ranges
            AttributeTextBox attText = new RangeTextBox("range1_comp", dispRange1, setRange1, "range");
            TextAttributeDic.Add(attText.attName, attText);

            attText = new RangeTextBox("range2_comp", dispRange2, setRange2, "range");
            TextAttributeDic.Add(attText.attName, attText);

            attText = new RangeTextBox("range1Comp_comp", dispR1CompCol, setR1CompCol, "column", false);
            TextAttributeDic.Add(attText.attName, attText);

            attText = new RangeTextBox("range2Comp_comp", dispR2CompCol, setR2CompCol, "column", false);
            TextAttributeDic.Add(attText.attName, attText);
            #endregion

            #region Settings
            
            
            attText = new AttributeTextBox("lowerTol_comp", dispLowerTol, true);
            attText.type = "double";
            attText.SetDefaultValue("0");
            TextAttributeDic.Add(attText.attName, attText);

            attText = new AttributeTextBox("upperTol_comp", dispUpperTol, true);
            attText.type = "double";
            attText.SetDefaultValue("0");
            TextAttributeDic.Add(attText.attName, attText);

            CustomAttribute customAtt = new CheckBoxAttribute("checkRangeSizes_comp", rangeSizeCheck, true);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            customAtt = new CheckBoxAttribute("resetFont_comp", resetFontCheck, false);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            customAtt = new CheckBoxAttribute("terminateRangeAtNull_comp", terminateAtNullCheck, false);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            customAtt = new CheckBoxAttribute("printDifference_comp", printDifferenceCheck, false);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            attText = new RangeTextBox("outputRange_comp", dispOutputRange, setOutputRange, "range");
            TextAttributeDic.Add(attText.attName, attText);

            customAtt = new CheckBoxAttribute("formatRange1_comp", formatRange1Check, false);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            customAtt = new CheckBoxAttribute("formatRange2_comp", formatRange2Check, false);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            customAtt = new CheckBoxAttribute("formatRangeOutput_comp", formatRangeOutputCheck, false);
            OtherAttributeDic.Add(customAtt.attName, customAtt);

            #endregion
        }

        CellFormatObject color1;
        CellFormatObject color2;
        private void CreateCellFormatObjects()
        {
            color1 = new CellFormatObject(sampleCell1, setFillColor1, resetFillColor1, setFontColor1, resetFontColor1,
                setBorderColor1, resetBorderColor1, getFormatFromCell1, applyFormat1);
            color1.ignoreDefaults = ignoreDefaultsCheck.Checked;

            color2 = new CellFormatObject(sampleCell2, setFillColor2, resetFillColor2, setFontColor2, resetFontColor2,
            setBorderColor2, resetBorderColor2, getFormatFromCell2, applyFormat2);
            color2.ignoreDefaults = ignoreDefaultsCheck.Checked;
        }

        private void AddToolTips()
        {
            #region Basic Comparison
            toolTip1.SetToolTip(rangeSizeCheck,
                "If unchecked, check will be done on the size of the smaller range.\n" +
                "For basic comparison: smaller row and column\n" +
                "For unique name comparison: smaller column number");

            #endregion

            #region Compare wiht Unique Name
            toolTip1.SetToolTip(compareRanges,
                "Cell to cell comparison.\n" +
                "Set font color to red for unequal values, gray for uncheckable.\n" +
                "Adjust settings below.");

            toolTip1.SetToolTip(setR1CompCol,
                "Column to be used as the unique name for comparing values for range 1");
            
            toolTip1.SetToolTip(setR2CompCol,
                "Column to be used as the unique name for comparing values for range 2");

            toolTip1.SetToolTip(compareWithUn,
                "Compares value based on unique name.\n" +
                "Set font color to red for unequal values, blue for unmatched values and gray for uncheckable values.\n" +
                "Uses range 1 and 2 from above.\n" +
                "Adjust settings below.");
            #endregion

            #region Settings
            toolTip1.SetToolTip(rangeSizeCheck,
                "Checks:\n" +
                "No. of rows and columns for basic comparison\n" +
                "No. of columns for unique name comparison");

            toolTip1.SetToolTip(dispLowerTol,
                "0.1 = 10%");
            
            toolTip1.SetToolTip(dispUpperTol,
                "0.1 = 10%");

            toolTip1.SetToolTip(toleranceLabel,
                "% difference is calculated based on abs[(range1 - range2)/range1]\n" +
                "Is +ve if range2 > range1" +
                "Cells with % difference that falls outside this range will be set to red font. Excludes exact boundary values.");

            toolTip1.SetToolTip(terminateAtNullCheck,
                "Resize ranges before proceeding. Will terminate range when the first row of empty values is encountered.\n" +
                "May be slow if the selected ranges are large");
            #endregion
        }

        #endregion

        #region Format Table

        private void formatTables_Click(object sender, EventArgs e)
        {
            Range selectedRange = thisApp.Selection;

            if (MessageBox.Show($"Warning: there is no way to undo this. Please backup your file before proceeding.\n\nProceed with formatting?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }

            try
            {
                // Define Range of cells based on selection
                int startRow = selectedRange.Row;
                int endRow = selectedRange.Rows.Count + selectedRange.Row - 1;
                int startCol = selectedRange.Column;
                int endCol = selectedRange.Columns.Count + selectedRange.Column - 1;

                //// Clear format of selected cells
                //if (!ignoreDefaultsCheck.Checked)
                //{
                //    selectedRange.Interior.ColorIndex = XlConstants.xlNone;
                //    selectedRange.Font.ColorIndex = XlConstants.xlAutomatic;
                //    selectedRange.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
                //}

                // Define start values
                Range compCol = ((RangeTextBox)TextAttributeDic["CompCol"]).GetRangeFromFullAddress();
                int compareColumnNum = compCol.Column;
                Worksheet selectedSheet = selectedRange.Worksheet;
                bool isColor1 = true;
                string currentValue = "";
                string nextValue = "";
                for (int currentRowNum = startRow; currentRowNum <= endRow; currentRowNum++)
                {
                    // Define Range
                    Range startCell = selectedSheet.Cells[currentRowNum, startCol];
                    Range endCell = selectedSheet.Cells[currentRowNum, endCol];
                    Range toFormatRange = selectedSheet.Range[startCell, endCell];

                    // Set Range color
                    if (isColor1)
                    {
                        color1.formatRange(toFormatRange);
                    }
                    else
                    {
                        color2.formatRange(toFormatRange);
                    }

                    // Check if we need to swap color for next iteration
                    if (currentRowNum == endRow) { continue; } // deal with final row
                    Range currentCell = selectedSheet.Cells[currentRowNum, compareColumnNum];
                    Range nextCell = selectedSheet.Cells[currentRowNum + 1, compareColumnNum];

                    if (ignoreBlanksCheck.Checked)
                    {

                        if (GetContentsAsString(currentCell) != "")
                        {
                            currentValue = currentCell.Value2.ToString(); // Update current value only if it is not null
                        }

                        if (GetContentsAsString(nextCell) != "")
                        {
                            nextValue = nextCell.Value2.ToString();
                        }
                        else
                        {
                            nextValue = currentValue; // Set to the same so it doesn't get updated
                        }
                    }
                    else
                    {
                        //currentValue = selectedSheet.Cells[currentRowNum, compareColumnNum].Value2.ToString();
                        //nextValue = selectedSheet.Cells[currentRowNum + 1, compareColumnNum].Value2.ToString();
                        currentValue = GetContentsAsString(currentCell);
                        nextValue = GetContentsAsString(nextCell);
                    }

                    if (currentValue != nextValue)
                    {
                        isColor1 = !isColor1;
                    }
                }
                MessageBox.Show("Complete formatting", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Warning: Error occurred, unable to complete formatting operation \n\n{ex.Message}", "Error");
            }
        }
        #endregion

        #region Insert Gap and Page Break
        int prevInput = 1;
        private void insertColGap_Click(object sender, EventArgs e)
        {
            #region Get User Input
            var userInput = thisApp.InputBox("Enter number of cells to insert ", "Enter data", prevInput, Type: 1);
            int numAddCells = 0;
            try
            {
                numAddCells = (int)userInput;
            }
            catch { return; }

            string msg = $"Confirm to add {numAddCells} columns?\n";
            if (checkShiftEntire.Checked)
            {
                msg += "Columns will be added.";
            }
            else
            {
                msg += "Cells will be shifted.";
            }

            if (MessageBox.Show(msg, "Confirmation", MessageBoxButtons.YesNo) == DialogResult.No) { return; }
            prevInput = numAddCells;
            #endregion

            Range selectedRange = thisApp.Selection;
            Worksheet thisSheet = selectedRange.Worksheet;
            (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(selectedRange);

            for (int colNum = endCol; colNum > startCol; colNum--)
            {
                if (checkShiftEntire.Checked)
                {
                    Range moveRange = thisSheet.Columns[colNum];
                    for (int repeat = 0; repeat < numAddCells; repeat++)
                    {
                        moveRange.Insert(XlInsertShiftDirection.xlShiftDown);
                    }
                }
                else
                {
                    Range startCell = thisSheet.Cells[startRow, colNum];
                    Range endCell = thisSheet.Cells[endRow, colNum];
                    Range moveRange = thisSheet.Range[startCell, endCell];
                    for (int repeat = 0; repeat < numAddCells; repeat++)
                    {
                        moveRange.Insert(XlInsertShiftDirection.xlShiftToRight);
                    }
                }

            }
        }

        private void insertRowGap_Click(object sender, EventArgs e)
        {
            #region Get User Input
            var userInput = thisApp.InputBox("Enter number of cells to insert ", "Enter data", prevInput, Type: 1);
            int numAddCells = 0;
            try
            {
                numAddCells = (int)userInput;
            }
            catch { return; }
            prevInput = numAddCells;
            #endregion

            Range selectedRange = thisApp.Selection;
            Worksheet thisSheet = selectedRange.Worksheet;
            (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(selectedRange);

            for (int rowNum = endRow; rowNum > startRow; rowNum--)
            {
                if (checkShiftEntire.Checked)
                {
                    Range moveRange = thisSheet.Rows[rowNum];
                    for (int repeat = 0; repeat < numAddCells; repeat++)
                    {
                        moveRange.Insert(XlInsertShiftDirection.xlShiftDown);
                    }
                }
                else
                {
                    Range startCell = thisSheet.Cells[rowNum, startCol];
                    Range endCell = thisSheet.Cells[rowNum, endCol];
                    Range moveRange = thisSheet.Range[startCell, endCell];
                    for (int repeat = 0; repeat < numAddCells; repeat++)
                    {
                        moveRange.Insert(XlInsertShiftDirection.xlShiftDown);
                    }
                }

            }
        }

        private void ignoreDefaultsCheck_CheckedChanged(object sender, EventArgs e)
        {
            color1.ignoreDefaults = ignoreDefaultsCheck.Checked;
            color2.ignoreDefaults = ignoreDefaultsCheck.Checked;
        }

        private void insertPageBreak_Click(object sender, EventArgs e)
        {
            #region Read Inputs
            int numRowOffset;
            int numColOffset;
            int numRepeats;

            try
            {
                numRowOffset = convertToInt(dispRowOffset.Text);
                numColOffset = convertToInt(dispColOffset.Text);
                numRepeats = convertToInt(dispNumRepeats.Text) - 1;
            }
            catch
            {
                return;
            }
            Range activeCell = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            Worksheet activeWorksheet = activeCell.Worksheet;
            int selectedRowNum = activeCell.Row;
            int selectedColNum = activeCell.Column;
            #endregion

            #region Apply page break
            int rowNum = selectedRowNum;
            int colNum = selectedColNum;
            if (numRowOffset > 0)
            {
                for (int iter = 0; iter <= numRepeats; iter++)
                {
                    rowNum += numRowOffset;
                    Range cell = activeWorksheet.Cells[rowNum, colNum];
                    if (rowNum > 1)
                    {
                        activeWorksheet.HPageBreaks.Add(cell);
                    }
                    if (numRowOffset == 0)
                    {
                        break;
                    }
                }
            }

            rowNum = selectedRowNum;
            if (numColOffset > 0)
            {
                for (int iter = 0; iter <= numRepeats; iter++)
                {
                    colNum += numColOffset;
                    Range cell = activeWorksheet.Cells[rowNum, colNum];
                    if (colNum > 1)
                    {
                        activeWorksheet.VPageBreaks.Add(cell);
                    }
                    if (numColOffset == 0)
                    {
                        break;
                    }
                }
            }


            //for (int rowIter = 0; rowIter <= numRepeats; rowIter++)
            //{
            //    rowNum = selectedRowNum + rowIter * numRowOffset;
            //    for (int colIter = 0; colIter <= numRepeats; colIter++)
            //    {
            //        colNum = selectedColNum + colIter * numColOffset;
            //        Range cell = activeWorksheet.Cells[rowNum, colNum];
            //        cell.Select();
            //        if (rowNum > 1)
            //        {
            //            activeWorksheet.HPageBreaks.Add(cell);
            //        }
            //        if (colNum > 1)
            //        {
            //            activeWorksheet.VPageBreaks.Add(cell);
            //        }

            //        if (numColOffset == 0)
            //        {
            //            break;
            //        }
            //    }
            //    if (numRowOffset == 0)
            //    {
            //        break;
            //    }
            //}
            #endregion

            int convertToInt(string text)
            {
                try
                {
                    if (text == "")
                    {
                        return 0;
                    }
                    else
                    {
                        return Convert.ToInt32(text);
                    }

                }
                catch
                {
                    throw new Exception($"Unable to convert '{text}' to number ");
                }
            }
        }

        #endregion

        #region Basic Comparison
        private void compareRanges_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                try
                {
                    #region Get Ranges
                    Range range1 = ((RangeTextBox)TextAttributeDic["range1_comp"]).GetRangeFromFullAddress();
                    Range range2 = ((RangeTextBox)TextAttributeDic["range2_comp"]).GetRangeFromFullAddress();
                    
                    #endregion

                    #region Reset Ranges
                    if (resetFontCheck.Checked)
                    {
                        progressTracker.UpdateStatus($"Checking ranges");
                        range1.Font.Color = Color.Black;
                        range2.Font.Color = Color.Black;
                    }
                    if (rangeSizeCheck.Checked) { AssertRangeSize(new Range[] { range1, range2 }, null, true); }
                    else { IntersectRanges(ref range1, ref range2); }
                    failedRanges = new List<Range>();
                    unmatchedRanges = new List<Range>();
                    #endregion

                    #region Compare Ranges
                    for (int rowNum = 1; rowNum <= range1.Rows.Count; rowNum++)
                    {
                        #region Update Progress
                        progressTracker.UpdateStatus($"Checking row {rowNum}");
                        worker.ReportProgress(ConvertToProgress(rowNum, range1.Rows.Count));
                        if (worker.CancellationPending) { return; }
                        #endregion

                        for (int colNum = 1; colNum <= range1.Columns.Count; colNum++)
                        {
                            Range cell1 = range1.Cells[rowNum, colNum];
                            Range cell2 = range2.Cells[rowNum, colNum];
                            CompareTwoCells(cell1, cell2);
                        }
                        
                    }
                    #endregion
                    
                    #region Format Ranges
                    progressTracker.UpdateStatus($"Formatting Ranges");
                    FormatRanges(failedRanges, Color.Red);
                    FormatRanges(unmatchedRanges, Color.Blue);
                    #endregion
                    MessageBox.Show($"Completed, number of failed ranges found: {failedRanges.Count / 2}", "Completed");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    DialogResult res = MessageBox.Show($"Comparison incomplete, format ranges that are compared?", "Confirmation", MessageBoxButtons.YesNo);
                    if (res == DialogResult.OK)
                    {
                        progressTracker.UpdateStatus($"Formatting Ranges");
                        FormatRanges(failedRanges, Color.Red);
                        FormatRanges(unmatchedRanges, Color.Blue);
                    }
                }
                finally
                {
                    failedRanges = null;
                    unmatchedRanges = null;
                }
            });
        }
        #endregion
        
        #region Compare with UN
        List<Range> failedRanges;
        List<Range> unmatchedRanges;
        private void compareWithUN_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                try
                {
                    #region Get Ranges
                    Range range1 = ((RangeTextBox)TextAttributeDic["range1_comp"]).GetRangeFromFullAddress();
                    Range range2 = ((RangeTextBox)TextAttributeDic["range2_comp"]).GetRangeFromFullAddress();
                    if (terminateAtNullCheck.Checked)
                    {
                        progressTracker.UpdateStatus($"Checking ranges");
                        TerminateRangeAtFirstNullRow(ref range1);
                        TerminateRangeAtFirstNullRow(ref range2);
                    }

                    if (rangeSizeCheck.Checked) { AssertRangeSize(new Range[] { range1, range2 }, "column", true); }
                    else { IntersectRanges(ref range1, ref range2, "column"); }

                    Worksheet sheet1 = range1.Worksheet;
                    Worksheet sheet2 = range2.Worksheet;
                    int range1CompColNum = (((RangeTextBox)TextAttributeDic["range1Comp_comp"]).GetRangeForSpecificSheet(sheet1.Name)).Column;
                    int range2CompColNum = (((RangeTextBox)TextAttributeDic["range2Comp_comp"]).GetRangeForSpecificSheet(sheet2.Name)).Column;
                    failedRanges = new List<Range>();
                    unmatchedRanges = new List<Range>();
                    #endregion

                    #region Reset Ranges
                    if (resetFontCheck.Checked)
                    {
                        range1.Font.Color = Color.Black;
                        range2.Font.Color = Color.Black;
                    }
                    #endregion

                    #region Create Hash Potato
                    Dictionary<string, List<Range>> range1Dict = AddRowToDictionary(range1, range1CompColNum);
                    Dictionary<string, List<Range>> range2Dict = AddRowToDictionary(range2, range2CompColNum);
                    #endregion

                    #region Compare Hash Potato
                    int completedItem = 0;
                    int maxItems = range1Dict.Keys.Count;
                    foreach (string uniqueName in range1Dict.Keys)
                    {
                        #region Update Progress
                        progressTracker.UpdateStatus($"Checking items for range 1: {completedItem+1}/{maxItems}");
                        worker.ReportProgress(ConvertToProgress(completedItem, maxItems));
                        if (worker.CancellationPending) { return; }
                        #endregion

                        List<Range> range1ranges = range1Dict[uniqueName];
                        List<Range> range2ranges;
                        if (!range2Dict.ContainsKey(uniqueName))
                        {
                            range2ranges = new List<Range>();
                        }
                        else
                        {
                            range2ranges = range2Dict[uniqueName];
                        }
                        CompareUniqueNameRange(range1ranges, range2ranges);
                        completedItem++;
                    }

                    // Deal with range2 keys that are not in range1
                    completedItem = 0;
                    maxItems = range2Dict.Keys.Count;
                    foreach (string uniqueName in range2Dict.Keys)
                    {
                        #region Update Progress
                        progressTracker.UpdateStatus($"Checking remaining items for range 2: {completedItem + 1}/{maxItems}");
                        worker.ReportProgress(ConvertToProgress(completedItem, maxItems));
                        if (worker.CancellationPending) { return; }
                        #endregion

                        if (range1Dict.ContainsKey(uniqueName))
                        {
                            continue; //Already dealth with
                        }
                        List<Range> range2ranges = range2Dict[uniqueName];
                        foreach (Range range in range2ranges)
                        {
                            //FormatUnmatchedRange(range);
                            unmatchedRanges.Add(range);
                            completedItem++;
                        }
                    }
                    #endregion

                    #region Format Ranges
                    progressTracker.UpdateStatus($"Formatting Ranges");
                    FormatRanges(failedRanges, Color.Red);
                    FormatRanges(unmatchedRanges, Color.Blue);
                    #endregion

                    MessageBox.Show($"Completed, number of failed ranges found: {failedRanges.Count/2}", "Completed");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    DialogResult res = MessageBox.Show($"Comparison incomplete, format ranges that are compared?", "Confirmation", MessageBoxButtons.YesNo);
                    if (res == DialogResult.OK)
                    {
                        progressTracker.UpdateStatus($"Formatting Ranges");
                        FormatRanges(failedRanges, Color.Red);
                        FormatRanges(unmatchedRanges, Color.Blue);
                    }
                }
                finally
                {
                    failedRanges = null;
                    unmatchedRanges = null;
                }
            });
        }

        #region Comparison Helpers
        private Dictionary<string, List<Range>> AddRowToDictionary(Range range, int rangeCompColNum)
        {
            Dictionary<string, List<Range>> rangeDict = new Dictionary<string, List<Range>>();
            Worksheet sheet = range.Worksheet;
            foreach (Range row in range.Rows) 
            {
                #region Get Unique Name
                int globalRowNum = row.Row;
                string uniqueName;
                if (sheet.Cells[globalRowNum, rangeCompColNum].Value2 == null)
                {
                    uniqueName = "";
                }
                else
                {
                    uniqueName = sheet.Cells[globalRowNum, rangeCompColNum].Value2.ToString();
                }
                
                #endregion

                #region Add to dictionary
                if (!rangeDict.ContainsKey(uniqueName))
                {
                    rangeDict[uniqueName] = new List<Range>();
                }
                rangeDict[uniqueName].Add(row);
                #endregion
            }

            return rangeDict;
        }

        private void CompareUniqueNameRange(List<Range> range1ranges, List<Range> range2ranges)
        {
            Range[] biggerRange;
            Range[] smallerRange;
            if (range1ranges.Count >= range2ranges.Count)
            {
                biggerRange = range1ranges.ToArray();
                smallerRange = range2ranges.ToArray();
            }
            else
            {
                biggerRange = range2ranges.ToArray();
                smallerRange = range1ranges.ToArray();
            }

            for (int i = 0; i < smallerRange.Length; i++)
            {
                Range rangeS = smallerRange[i];
                Range rangeB = biggerRange[i];
                CompareRows(rangeB, rangeS);
            }

            // Deal with remaining bigger ranges
            if (biggerRange.Length != smallerRange.Length)
            {
                for (int i = smallerRange.Length; i < biggerRange.Length; i++)
                {
                    Range row = biggerRange[i];
                    //FormatUnmatchedRange(row);
                    unmatchedRanges.Add(row);
                }
            }
        }

        private void CompareRows(Range range1, Range range2)
        {
            for (int colNum = 1; colNum <= range1.Columns.Count; colNum++)
            {
                Range cell1 = range1.Cells[colNum];
                Range cell2 = range2.Cells[colNum];
                CompareTwoCells(cell1, cell2);                
            }
        }

        private void CompareTwoCells(Range cell1, Range cell2)
        {
            var cell1Val = cell1.Value2;
            var cell2Val = cell2.Value2;

            if (cell1Val == null || cell2Val == null) { return; }
            bool cell1Check = double.TryParse(cell1Val.ToString(), out double cell1Double);
            bool cell2Check = double.TryParse(cell2Val.ToString(), out double cell2Double);

            #region Not Double
            if (!cell1Check || !cell2Check )
            {
                if (cell1Val == cell2Val) { return; }
                else
                {
                    failedRanges.Add(cell1);
                    failedRanges.Add(cell2);
                    //FormatFailedRange(cell1);
                    //FormatFailedRange(cell2);
                    return;
                }
            }
            #endregion

            #region Compare Values

            if (cell1Double == cell2Double) { return; }

            double lowerBound = TextAttributeDic["lowerTol_comp"].GetDoubleFromTextBox();
            double upperBound = TextAttributeDic["upperTol_comp"].GetDoubleFromTextBox();
            
            double diff;
            if (cell1Double == 0) 
            {
                if (cell2Double > 0) { diff = double.PositiveInfinity; }
                else { diff = double.NegativeInfinity; }
            }
            else 
            {
                diff = Math.Abs(cell1Double - cell2Double) / Math.Abs(cell1Double);
                if (cell1Double - cell2Double > 0)
                {
                    diff = -diff;
                }
            }
            
            if (diff < lowerBound || diff > upperBound)
            {
                //FormatFailedRange(cell1);
                //FormatFailedRange(cell2);
                failedRanges.Add(cell1);
                failedRanges.Add(cell2);
            }
            #endregion
        }

        private void FormatRanges(List<Range> ranges, Color color)
        {
            foreach (Range range in ranges) { range.Font.Color = color; }
        }

        #region Old
        //private void FormatFailedRange()
        //{
        //    //range.Interior.Color = Color.LightPink;
        //    //range.Font.Color = Color.Red;
        //}
        //private void FormatUnmatchedRange()
        //{
        //    //range.Interior.Color = Color.PaleTurquoise;
        //    //range.Font.Color = Color.Blue;
        //}

        //private void FormatErrorRange(Range range)
        //{
        //    range.Font.Color = Color.Gray;
        //}
        #endregion
        #endregion

        #endregion
    }
}

