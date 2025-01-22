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
using System.IO;
using ExcelAddIn2.Excel_Pane_Folder;
using static ExcelAddIn2.CommonUtilities;
using ExcelAddIn2.Piling;
using TextBox = System.Windows.Forms.TextBox;


namespace ExcelAddIn2
{
    public partial class IterationPane : UserControl
    {
        #region Initialisers
        Workbook ThisWorkBook;
        Microsoft.Office.Interop.Excel.Application ThisApplication;
        DocumentProperties AllCustProps;
        Dictionary<string, AttributeTextBox> RangeAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();

        public IterationPane()
        {
            InitializeComponent();
            ThisApplication = Globals.ThisAddIn.Application;
            ThisWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            AllCustProps = ThisWorkBook.CustomDocumentProperties;
            CreateAttributes();
            AddToolTips();
            AddHeaders();
        }

        private void AddHeaders()
        {
            string[,] headerRange = new string[,]
            {
                {"Iteration Name", "Status", "Input Att", "Output Att"},
                {"","","A1","~A2" }
            };
            AddHeaderMenuToButton(SetHeadR1, headerRange);

            List<string> headers = new List<string>
            {
                "Name",
                "Iteration Data Source",
                "Destination Columns",
                "Criteria Source Col",
                "Criteria logic",
                "Criteria Value",
                "Status Col",
                "Iteration Mode",
                "Optimisation Source Column",
                "Optimisation Mode",
                "Optimisation Target Value (if applicable)"
            };
            AddHeaderMenuToButton(SetItDataTable, headers);
        }

        private void CreateAttributes()
        {
            // Create Attribute Objects 
            #region Setup Tools
            ComboBoxAttribute FormatType = new ComboBoxAttribute("FormatType", FormatOptions, "3 Set fill and font");
            OtherAttributeDic.Add("FormatType", FormatType); 
            #endregion

            #region Utilities
            MultipleSheetsAttribute KeepSheets = new MultipleSheetsAttribute("KeepSheets", DelSheets, true);
            OtherAttributeDic.Add("KeepSheets", KeepSheets); 
            MultipleSheetsAttribute SavedDupeSheet = new MultipleSheetsAttribute("SavedDupeSheet", SetDupeSheet);
            OtherAttributeDic.Add("SavedDupeSheet", SavedDupeSheet); 
            #endregion

            #region Calculation Settings (Universal)
            RangeTextBox HeadRange1 = new RangeTextBox("HeadRange1", DispHeadR1, SetHeadR1, "row");
            RangeAttributeDic.Add("HeadRange1", HeadRange1);
            SheetTextBox OutSheet1 = new SheetTextBox("OutSheet1", DispOutS1, SetOutSheet1);
            RangeAttributeDic.Add("OutSheet1", OutSheet1);
            RangeTextBox InRange1 = new RangeTextBox("InRange1", DispInputR1, SetInRange1, "column");
            RangeAttributeDic.Add("InRange1", InRange1);
            #endregion

            #region Multiple Runs Iteration Setting
            RangeTextBox ItSource = new RangeTextBox("ItSource", DispItSource, SetItSource, "range");
            RangeAttributeDic.Add("ItSource", ItSource);
            RangeTextBox ItDest = new RangeTextBox("ItDest", DispItDest, SetItDest, "row");
            RangeAttributeDic.Add("ItDest", ItDest);
            RangeTextBox CriteriaSource = new RangeTextBox("CriteriaSource", DispCriteriaSource, SetCriteriaSource, "column");
            RangeAttributeDic.Add("CriteriaSource", CriteriaSource);
            RangeTextBox StatusCol = new RangeTextBox("StatusCol", DispStatusCol, SetStatusCol, "column");
            RangeAttributeDic.Add("StatusCol", StatusCol);
            ComboBoxAttribute logicSymbol = new ComboBoxAttribute("logicSymbol", DispLogicSymbol, "<=");
            OtherAttributeDic.Add("logicSymbol", logicSymbol);
            AttributeTextBox CriteriaValue = new AttributeTextBox("CriteriaValue", DispCriteriaValue, true);
            RangeAttributeDic.Add("CriteriaValue", CriteriaValue);
            ComboBoxAttribute iterMode = new ComboBoxAttribute("iterMode", DispIterationMode, "1 Stop when condition met");
            OtherAttributeDic.Add("iterMode", iterMode);
            #endregion

            #region Optimisation Settings
            //Optmise values
            RangeTextBox OptimiseCol = new RangeTextBox("OptimiseCol", DispOptimiseCol, SetOptimiseCol, "column");
            RangeAttributeDic.Add("OptimiseCol", OptimiseCol);
            ComboBoxAttribute OptimiseType = new ComboBoxAttribute("OptimiseType", DispOptimiseType, "1 Minimise");
            OtherAttributeDic.Add("OptimiseType", OptimiseType);
            AttributeTextBox OptimiseTarget = new AttributeTextBox("OptimiseTarget", DispOptimiseTarget, true);
            RangeAttributeDic.Add("OptimiseTarget", OptimiseTarget);
            #endregion

            #region Multiple Criteria Iteration
            RangeTextBox ItTable = new RangeTextBox("ItTable", DispItDataTable, SetItDataTable, "range");
            RangeAttributeDic.Add("ItTable", ItTable);
            #endregion

            #region Sheet Tools
            SheetTextBox CopySheet = new SheetTextBox("CopySheet", DispCopySheet, SetCopySheet);
            RangeAttributeDic.Add("CopySheet", CopySheet);
            #endregion

            #region Misc

            #endregion

            #region Single Cell Iteration 
            RangeTextBox CriteriaSource2 = new RangeTextBox("CriteriaSource2", dispCriteriaSource2, setCriteriaSource2, "column", false);
            RangeAttributeDic.Add("CriteriaSource2", CriteriaSource2);
            ComboBoxAttribute LogicSymbol2 = new ComboBoxAttribute("LogicSymbol2", dispLogicSymbol2, "<=");
            OtherAttributeDic.Add("LogicSymbol2", LogicSymbol2);
            AttributeTextBox CriteriaValue2 = new AttributeTextBox("CriteriaValue2", dispCriteriaValue2, true);
            RangeAttributeDic.Add("CriteriaValue2", CriteriaValue2);

            AttributeTextBox Increment = new AttributeTextBox("Increment", dispIncrement, true);
            Increment.type = "double";
            Increment.textBox.Text = "0.1";
            RangeAttributeDic.Add("Increment", Increment);

            AttributeTextBox LoopNum = new AttributeTextBox("LoopNum", dispLoopNum, true);
            LoopNum.type = "int";
            LoopNum.textBox.Text = "100";
            RangeAttributeDic.Add("LoopNum", LoopNum);
            #endregion
        }

        private void AddToolTips()
        {
            #region Print Tools
            toolTip1.SetToolTip(SetCopySheet,
                "Set sheet to duplicate");
            toolTip1.SetToolTip(duplicateSheets,
                "Duplicate the sheet defined on the left.\n" +
                "Number of sheets and name of output sheet is read from current range selection.");
            #endregion

        }
        #endregion


        #region Setup Tools

        #region Helper Functions
        #region Input Checkers
        private (bool, string, string) CheckIfRangeHasSheet(string InputAddress)
        {
            string[] parts = InputAddress.Split('!');
            if (parts.Length == 2)
            {
                return (true, parts[0], parts[1]);
            }
            else
            {
                return (false, DispOutS1.Text, InputAddress);
            }
        }

        private bool CheckIfSheetIsValid(string SheetName, bool errorMsg = false)
        {
            try
            {
                Worksheet worksheet = ThisWorkBook.Sheets[SheetName];
                return true;
            }
            catch (Exception)
            {
                if (errorMsg)
                {
                    MessageBox.Show($"Sheet '{SheetName}' does not exist.");
                }
                return false;
            }
        }

        private bool CheckIfAddressIsValid(string CellAddress, bool errorMsg = false)
        {
            // Check if valid worksheet and range
            try
            {
                Worksheet worksheet = ThisWorkBook.ActiveSheet;
                Range excelRange = worksheet.Range[CellAddress];
                return true;
            }
            catch (Exception)
            {
                if (errorMsg)
                {
                    MessageBox.Show($"Cell '{CellAddress}' does not exist.");
                }
                return false;
            }
        }

        private bool CheckIfInputIsValid(string InputAddress, bool errorMsg = false)
        {
            (bool hasSheet, string SheetName, string CellAddress) = CheckIfRangeHasSheet(InputAddress);
            bool isSheetValid = true;
            if (hasSheet)
            {
                isSheetValid = CheckIfSheetIsValid(SheetName, errorMsg);
            }
            bool isAddValid = CheckIfAddressIsValid(CellAddress, errorMsg);
            if (isSheetValid && isAddValid)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private Range GetRangeFromFullAddress(string FullAddress)
        {
            var parts = FullAddress.Split('!');

            if (parts.Length != 2)
                throw new ArgumentException("Invalid address format. Expected format: SheetName!CellAddress");

            try
            {
                //ThisWorkBook.Sheets[parts[0]].Activate();
                Worksheet ThisWorksheet = ThisWorkBook.Sheets[parts[0]];
                Range returnRange = ThisWorksheet.Range[parts[1]];
                return returnRange;
            }
            catch
            {
                MessageBox.Show($"Error Returning Range at ${FullAddress}");
                return null;
            }
        }

        private Range GetRangeFromAllAddress(string InputAddress)
        {
            (bool hasSheet, string SheetName, string CellAddress) = CheckIfRangeHasSheet(InputAddress);
            return GetRangeFromFullAddress(SheetName + "!" + CellAddress);
        }
        #endregion

        #region Set Linked Colour
        private void SetLinkedCellFormat(string InputCellSheetAddress, string OutputCellSheetAddress, int type)
        {
            // "type" indidcates type of changes to be made
            // 1 set fill colour
            // 2 set font colour
            // 3 set fill and font
            // 4 reset fill 
            // 5 reset font
            // 6 reset fill and font

            Range CurrentCell = GetRangeFromFullAddress(InputCellSheetAddress);
            Range TargetCell = GetRangeFromFullAddress(OutputCellSheetAddress);

            if (CurrentCell == null | TargetCell == null)
            {
                MessageBox.Show("Invalid input cells");
                return;
            }
            bool completed = true;

            // Set Fill Colour
            if (type == 1 | type == 3)
            {
                try
                {
                    //TargetCell.Interior.Color = CurrentCell.Interior.Color;
                    TargetCell.Interior.ColorIndex = CurrentCell.Interior.ColorIndex;
                }
                catch
                {
                    completed = false;
                }
            }

            // Set Font Colour
            if (type == 2 | type == 3)
            {
                try
                {
                    TargetCell.Font.Color = CurrentCell.Font.Color;
                }
                catch
                {
                    completed = false;
                }
            }

            // Reset Fill 
            if (type == 4 | type == 6)
            {
                try
                {
                    TargetCell.Interior.ColorIndex = 0;
                }
                catch
                {
                    //MessageBox.Show($"Failed to set cell colour for {OutputCellSheetAddress}.");
                    completed = false;
                }
            }

            // Reset Font 
            if (type == 5 | type == 6)
            {
                try
                {
                    TargetCell.Font.Color = System.Drawing.Color.Black.ToArgb();
                }
                catch
                {
                    //MessageBox.Show($"Failed to set cell colour for {OutputCellSheetAddress}.");
                    completed = false;
                }
            }

            if (!completed)
            {
                MessageBox.Show($"Failed to set format for {OutputCellSheetAddress}.");
            }
        }
        #endregion

        #endregion

        private void SetCellName_Click(object sender, EventArgs e)
        {
            #region Archive 
            ////Select output range
            //Range CurrentSelection;
            //try
            //{
            //    CurrentSelection = Globals.ThisAddIn.Application.InputBox("Select cell(s) to overwrite.", "Select cell.", Type: 8);

            //}
            //catch
            //{
            //    return;
            //}
            //Worksheet CurrentSheet = ThisWorkBook.Sheets[CurrentSelection.Worksheet.Name];
            #endregion
            Range CurrentSelection = ThisApplication.ActiveWindow.RangeSelection;
            Worksheet CurrentSheet = ThisApplication.ActiveSheet;
            
            // Convert row and column from current selection into index
            List<Range> ListSelection = new List<Range>();
            foreach (Range CurrentCell in CurrentSelection.Cells)
            {
                ListSelection.Add(CurrentCell);
            }
            int outputCellNum = 0;
            while (outputCellNum < CurrentSelection.Cells.Count)
            {
                Range TargetRanges;
                try
                {
                    Range CurrentCell = ListSelection[outputCellNum];
                    Range LabelCell = CurrentCell.Offset[-1, 0];
                    string msg;
                    if (LabelCell.Value2 != null)
                    {
                        msg = $"Select target cell(s) for '{LabelCell.Value2}'.";
                    }
                    else
                    {
                        msg = $"Select target cell(s).";
                    }

                    TargetRanges = Globals.ThisAddIn.Application.InputBox(msg, "Select cell.", Type: 8);
                }
                catch (Exception)
                {
                    return;
                }

                foreach (Range TargetRange in TargetRanges.Cells)
                {
                    if (outputCellNum == CurrentSelection.Count)
                    {
                        MessageBox.Show("Number of cells in selection exceeds number of cells for input. Remaining cells excluded.");
                        break;
                    }
                    string PrintValue = "";
                    if (IsOutputCheck.Checked)
                    {
                        PrintValue += "~";
                    }
                    if (SetShtNmCheck.Checked)
                    {
                        PrintValue += TargetRange.Worksheet.Name + "!" + TargetRange.get_Address(false, false);
                    }
                    else
                    {
                        PrintValue += TargetRange.get_Address(false, false);
                    }
                    Range CurrentCell = ListSelection[outputCellNum];
                    CurrentCell.Value2 = PrintValue;
                    outputCellNum++;
                }
                ThisWorkBook.Sheets[TargetRanges.Worksheet.Name].Activate();
                TargetRanges.Select();
            }
            CurrentSheet.Activate();
        }

        private void SetRangeName_Click(object sender, EventArgs e)
        {
            Range CurrentSelection = ThisApplication.ActiveWindow.RangeSelection;
            Worksheet CurrentSheet = ThisApplication.ActiveSheet;

            // Convert row and column from current selection into index
            List<Range> ListSelection = new List<Range>();
            foreach (Range CurrentCell in CurrentSelection.Cells)
            {
                ListSelection.Add(CurrentCell);
            }
            
            int outputCellNum = 0; // Count of input for user selected range
            while (outputCellNum < CurrentSelection.Cells.Count)
            {
                #region Get user to select range
                Range TargetRange;
                Range CurrentCell;
                try
                {
                    CurrentCell = ListSelection[outputCellNum];
                    Range LabelCell = CurrentCell.Offset[-1, 0]; // This offset has to be updated
                    string msg;
                    if (LabelCell.Value2 != null)
                    {
                        msg = $"Select target range for '{LabelCell.Value2}'.";
                    }
                    else
                    {
                        msg = $"Select target range.";
                    }

                    TargetRange = Globals.ThisAddIn.Application.InputBox(msg, "Select range.", Type: 8);
                }
                catch (Exception)
                {
                    return;
                }
                #endregion

                #region Output value into original cells
                string PrintValue = "";
                if (IsOutputCheck.Checked)
                {
                    PrintValue += "~";
                }
                if (SetShtNmCheck.Checked)
                {
                    PrintValue += TargetRange.Worksheet.Name + "!" + TargetRange.get_Address(false, false);
                }
                else
                {
                    PrintValue += TargetRange.get_Address(false, false);
                }

                CurrentCell.Value2 = PrintValue;
                #endregion

                outputCellNum++;
            }
            CurrentSheet.Activate();
        }

        private void ShowOutputCell_Click(object sender, EventArgs e)
        {
            Worksheet CurrentSheet = Globals.ThisAddIn.Application.ActiveWindow.ActiveSheet;

            Range CurrentSelection = ThisApplication.ActiveWindow.RangeSelection;
            foreach (Range CurrentCell in CurrentSelection.Cells)
            {
                //Range CurrentCell = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                string InputAddress = CurrentCell.Text;
                if (InputAddress == "")
                {
                    continue;
                }
                if (InputAddress[0] == '~')
                {
                    InputAddress = InputAddress.Substring(1);
                }

                // Check if input is valid
                if (InputAddress == null)
                {
                    return;
                }
                bool isValid = CheckIfInputIsValid(InputAddress, true);
                if (!isValid)
                {
                    return;
                }

                // Finalise input info
                (bool hasSheet, string SheetName, string CellAddress) = CheckIfRangeHasSheet(InputAddress);
                if (!hasSheet)
                {
                    SheetName = DispOutS1.Text; // Take sheet name from variable set
                }

                // Select Sheet
                Worksheet TargetWorksheet;
                try
                {
                    TargetWorksheet = ThisWorkBook.Sheets[SheetName];
                    TargetWorksheet.Activate();
                }
                catch (Exception)
                {
                    MessageBox.Show("Error finding sheet");
                    return;
                }

                // Try Selecting Cell
                try
                {
                    Range TargetCell = TargetWorksheet.Range[CellAddress];
                    Range LabelCell = CurrentCell.Offset[-1, 0];
                    TargetCell.Select();
                    if (ReturnCheck.Checked)
                    {
                        if (LabelCell.Value2 != null)
                        {
                            MessageBox.Show($"Cell labelled '{LabelCell.Value2}' with content '{InputAddress}' selected. Click ok to continue/return.");
                        }
                        else
                        {
                            MessageBox.Show($"Cell with content '{InputAddress}' selected. Click ok to return.");
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Cell labelled '{LabelCell.Value2}' with content '{InputAddress}' selected. Click ok to continue.");
                    }

                }
                catch (Exception)
                {
                    MessageBox.Show("Error selecting cell.");
                    return;
                }
            }
            // Go back to original cell
            if (ReturnCheck.Checked)
            {
                CurrentSheet.Activate();
                CurrentSelection.Select();
            }
        }

        private void FormatLinkCell_Click(object sender, EventArgs e)
        {
            int type;
            // Check if FormatOptions (drop down menu) is filled
            try
            {
                object checkFormat = FormatOptions.SelectedItem;
                type = FormatOptions.SelectedItem.ToString()[0] - '0';
            }
            catch
            {
                MessageBox.Show("Select format option.");
                return;
            }
            
            // Format Selected Cells
            Range CurrentSelection = ThisApplication.ActiveWindow.RangeSelection;
            int numCellsUpdated = 0;
            foreach (Range cell in CurrentSelection.Cells)
            {
                string InputCellSheetAddress = cell.Worksheet.Name + "!" + cell.Address[false, false];
                string OutputCellSheetAddress = cell.Text;
                if (OutputCellSheetAddress == "")
                {
                    continue;
                }

                bool isValid = CheckIfInputIsValid(OutputCellSheetAddress, true);
                if (!isValid)
                {
                    continue;
                }

                (bool hasSheet, string SheetName, string CellAddress) = CheckIfRangeHasSheet(OutputCellSheetAddress);
                OutputCellSheetAddress = SheetName + "!" + CellAddress;
                SetLinkedCellFormat(InputCellSheetAddress, OutputCellSheetAddress, type);
                numCellsUpdated += 1;
            }
            MessageBox.Show($"{numCellsUpdated}/{CurrentSelection.Count} cells format updated.","Completed");
        }
        #endregion


        #region Utilities
        private void ClearOutput_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Confirm to delete data in all output region? This cannot be undone.", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            #region Get Inputs
            (Range headerRange, Range InputRange, Worksheet OutputSheet) = GetSingleRunInputs();
            if (headerRange == null)
            {
                return;
            }

            // Convert header into list
            (List<(Range, string)> OutputHeaders, List<(Range, string)> InputHeaders) = ConvertHeaders(headerRange);
            if (OutputHeaders.Count == 0)
            {
                return;
            }
            #endregion

            // Delete
            foreach (Range Row in InputRange.Rows)
            {
                int ExcelRowNum = Row.Row;
                foreach ((Range HeaderCell, string TargetAddress) in OutputHeaders)
                {
                    try
                    {
                        //HeaderCell.Worksheet.Select();
                        Range sourceRange = HeaderCell.Worksheet.Cells[ExcelRowNum, HeaderCell.Column];
                        sourceRange.ClearContents();
                        continue;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Unable to clear target cell(s)\n" + ex, "Error");
                        return;
                    }
                }

            }
        }

        private void clearIterationInputs_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Comfirm to clear user inputs from excel pane? \nThis cannot be undone.", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            foreach (KeyValuePair<string, CustomAttribute> attribute in OtherAttributeDic)
            {
                attribute.Value.ResetValue();
            }

            foreach (KeyValuePair<string, AttributeTextBox> attribute in RangeAttributeDic)
            {
                attribute.Value.ResetValue();
            }
        }

        private void ExportUserInputs_Click(object sender, EventArgs e)
        {
            List<object> propName = new List<object>();
            List<object> propValue = new List<object>();

            foreach (DocumentProperty prop in AllCustProps)
            {
                if (!(OtherAttributeDic.Keys.Contains(prop.Name) || RangeAttributeDic.Keys.Contains(prop.Name)))
                {
                    continue; // Skip properties not in this excel pane
                }
                propName.Add(prop.Name);
                propValue.Add(prop.Value.ToString());
            }
            DialogResult result = MessageBox.Show("Confirm to export value? This will override cell values at current selection and cannot be undone.\n" +
                "Output table size:\n" +
                $"Number of rows: {propName.Count}\n" +
                "Number of columns: 2", "Confirmation");
            WriteListToExcel(0, 0, propName, propValue);
        }

        private void ImportUserInputs_Click(object sender, EventArgs e)
        {
            #region Confirmation
            DialogResult result = MessageBox.Show("Comfirm to read user inputs from selection? \nThis cannot be undone.", "Confirmation", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }
            #endregion

            #region Checks
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            if (selectedRange.Columns.Count != 2)
            {
                MessageBox.Show("Incorrect number of columns selected");
                return;
            }
            #endregion

            List<string> propertyFailedToAdd = new List<string>();
            foreach (Range row in selectedRange.Rows)
            {
                string propName = row.Cells[1][1].Value2.ToString();
                bool success = false;

                if (row.Cells[2][1].Value2 == null)
                {
                    propertyFailedToAdd.Add(propName);
                    continue;
                }
                string propValue = row.Cells[2][1].Value2.ToString();
                
                if (OtherAttributeDic.Keys.Contains(propName))
                {
                    success = OtherAttributeDic[propName].ImportValue(propValue);
                }
                else if (RangeAttributeDic.Keys.Contains(propName))
                {
                    success = RangeAttributeDic[propName].ImportValue(propValue);
                }

                if (!success)
                {
                    propertyFailedToAdd.Add(propName);
                }
            }

            #region Termination Message
            if (propertyFailedToAdd.Count > 0 )
            {
                string msg = "Unable to import property for the following properties:";
                foreach (string prop in propertyFailedToAdd)
                {
                    msg += "\n" + prop;
                }
                MessageBox.Show(msg, "Warning");
            }
            else
            {
                MessageBox.Show("Import Completed.", "Complete");
            }
            #endregion
        }
        #region Helper Functions
        //private void WriteListToExcel(int rowOff, int colOff, List<List<object>> listOfList)
        private void WriteListToExcel(int rowOff, int colOff, params List<object>[] listOfLists)
        {
            // This code takes any number of list (of various types) and outputs them into excel 
            // Output order depends on order of the input array
            // Output location is the first cell of the current selection, offset by rowOff and colOff

            // Find number of rows and columns
            int numRow = 0;
            int numCol = listOfLists.Count();
            foreach (List<object> thisList in listOfLists)
            {
                if (thisList.Count() > numRow)
                {
                    numRow = thisList.Count(); // Finds max number of rows out of all the various arrays
                }
            }

            // Initiate object
            object[,] dataArray = new object[numRow, numCol];
            int col = 0;
            foreach (List<object> thisList in listOfLists)
            {
                int row = 0;
                foreach (object entry in thisList)
                {
                    dataArray[row, col] = entry.ToString();
                    row += 1;
                }
                col += 1;
            }

            // Add section to read input data from Excel
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            // Write to Excel
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Range startCell = activeSheet.Cells[selectedRange.Row + rowOff, selectedRange.Column + colOff];
            Range endCell = startCell.Offset[numRow - 1, numCol - 1];
            Range writeRange = activeSheet.Range[startCell, endCell];
            writeRange.Value2 = dataArray;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        #endregion
        #endregion


        #region Run Single

        #region Helper Functions
        private (Range, Range, Worksheet) GetSingleRunInputs()
        {
            // Header Range
            Range headerRange;
            try
            {
                headerRange = GetRangeFromFullAddress(DispHeadR1.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for header.");
                return (null, null, null);
            }
            if (headerRange.Rows.Count != 1)
            {
                MessageBox.Show("Invalid input for header. Only one row allowed.");
                return (null, null, null);
            }

            // Input Range
            Range InputRange;
            try
            {
                if (OverrideInputCheck.Checked)
                {
                    InputRange = ThisApplication.ActiveWindow.Selection;
                    InputRange = InputRange.Columns[1];
                }
                else
                {
                    InputRange = GetRangeFromFullAddress(DispInputR1.Text);
                }
            }
            catch
            {
                MessageBox.Show("Invalid input for input range.");
                return (null, null, null);
            }

            // Output Sheet
            Worksheet OutputSheet;
            try
            {
                OutputSheet = ThisWorkBook.Sheets[DispOutS1.Text];
            }
            catch
            {
                MessageBox.Show("Invalid input for output sheet.");
                return (null, null, null);
            }

            return (headerRange, InputRange, OutputSheet);
        }

        private (List<(Range, string)>, List<(Range, string)>) ConvertHeaders(Range headerRange)
        {
            // All headers should be converted to "sheet name"!"Cell address"
            //List<Range> InputHeaders = new List<Range>();
            List<(Range, string)> OutputHeaders = new List<(Range, string)>();
            List<(Range, string)> InputHeaders = new List<(Range, string)>();
            foreach (Range HeaderCell in headerRange)
            {
                if (HeaderCell.Text.Length == 0)
                {
                    continue; // skip blank cells
                }
                if (HeaderCell.Text[0] == '~') // is output
                {
                    // Check that range can be accessed
                    Range TargetCell = GetRangeFromAllAddress(HeaderCell.Text.Substring(1));
                    if (TargetCell == null)
                    {
                        return (new List<(Range, string)>(), new List<(Range, string)>());
                    }
                    // Add to list
                    string InputAddress = HeaderCell.Text.Substring(1);
                    (bool hasSheet, string SheetName, string CellAddress) = CheckIfRangeHasSheet(InputAddress);
                    OutputHeaders.Add((HeaderCell, SheetName + "!" + CellAddress));
                }
                else // is input
                {
                    // Check that range can be accessed
                    Range TargetCell = GetRangeFromAllAddress(HeaderCell.Text);
                    if (TargetCell == null)
                    {
                        return (new List<(Range, string)>(), new List<(Range, string)>());
                    }
                    // Add to list
                    string InputAddress = HeaderCell.Text;
                    (bool hasSheet, string SheetName, string CellAddress) = CheckIfRangeHasSheet(InputAddress);
                    InputHeaders.Add((HeaderCell, SheetName + "!" + CellAddress));
                }
            }
            return (OutputHeaders, InputHeaders);
        }

        private bool GetandSetTrueValues(
            List<(Range, string)> OutputHeaders, 
            List<(Range, string)> InputHeaders,
            int RowstoRunRowNum,
            List<Worksheet> OGSheets)
        {
            ThisApplication.Calculation = XlCalculation.xlCalculationAutomatic;
            // Set values
            foreach ((Range HeaderCell, string TargetAddress) in InputHeaders)
            {
                string SourceAddress = "";
                try
                {
                    Range sourceRange = HeaderCell.Worksheet.Cells[RowstoRunRowNum, HeaderCell.Column];
                    SourceAddress = sourceRange.Worksheet.Name + "!" + sourceRange.Address[false, false];
                    if (sourceRange.Text != "*NC")
                    {
                        string checktext = sourceRange.Text;
                        Range targetRange = GetRangeFromAllAddress(TargetAddress);
                        targetRange.Value2 = sourceRange.Value2;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unable to update target cell {TargetAddress} from source cell {SourceAddress}. \n" + ex);
                    ResetOGSheets(OGSheets);
                    return false;
                }
            }

            // Get values
            foreach ((Range HeaderCell, string TargetAddress) in OutputHeaders)
            {
                ThisApplication.Calculate();
                string SourceAddress = "";
                try
                {
                    Range sourceRange = HeaderCell.Worksheet.Cells[RowstoRunRowNum, HeaderCell.Column];
                    SourceAddress = sourceRange.Worksheet.Name + "!" + sourceRange.Address[false, false];
                    if (sourceRange.Text != "*NC")
                    {
                        string checktext = sourceRange.Text;
                        Range targetRange = GetRangeFromAllAddress(TargetAddress);
                        sourceRange.Worksheet.Activate();
                        sourceRange.Value2 = targetRange.Value2;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unable to update target cell {SourceAddress} from source cell {TargetAddress}. \n" + ex);
                    ResetOGSheets(OGSheets);
                    return false;
                }
            }
            return true;
        }

        #region ManageSheets
        private Worksheet CopyNewSheet(Worksheet sourceSheet, string newName)
        {
            sourceSheet.Copy(After: sourceSheet);
            Worksheet backUpSheet = ThisWorkBook.Sheets[sourceSheet.Index + 1];
            //Check if sheet with same name exist, if so, delete
            try
            {
                ThisApplication.DisplayAlerts = false;

                Worksheet checkSheet = ThisWorkBook.Sheets[newName];
                checkSheet.Delete();
            }
            catch (System.Runtime.InteropServices.COMException) { }
            finally
            {
                ThisApplication.DisplayAlerts = true;
            }
            backUpSheet.Name = newName;
            return backUpSheet;
        }

        private void ResetOGSheets(List<Worksheet> OGSheets)
        {
            try
            {
                //ThisApplication.DisplayAlerts = false;
                foreach (Worksheet OGWorksheet in OGSheets)
                {
                    string newName = OGWorksheet.Name.Substring(0, OGWorksheet.Name.Length - 3);
                    Worksheet backUpSheet = CopyNewSheet(OGWorksheet, newName);
                    ThisApplication.DisplayAlerts = false;
                    OGWorksheet.Delete();
                }
            }
            catch { }
            finally
            {
                ThisApplication.ScreenUpdating = true;
                ThisApplication.DisplayAlerts = true;
            }
        }

        private List<Worksheet> RenameBaseSheets()
        {
            List<Worksheet> OGSheets = new List<Worksheet>();
            HashSet<string> PrintSheets;
            if (CheckNewSheet1.Checked)
            {
                try
                {
                    PrintSheets = ((MultipleSheetsAttribute)OtherAttributeDic["SavedDupeSheet"]).GetSheetNamesHash();
                    if (PrintSheets == null || PrintSheets.Count == 0)
                    {
                        MessageBox.Show("Set sheets to duplicate");
                        return OGSheets;
                    }
                }
                catch
                {
                    MessageBox.Show("Set sheets to duplicate");
                    return OGSheets;
                }

                foreach (string sheet in PrintSheets)
                {
                    string newName = sheet + "_OG";
                    Worksheet backUpSheet = CopyNewSheet(ThisWorkBook.Sheets[sheet], newName);
                    OGSheets.Add(backUpSheet);
                }
            }
            else
            {
                string sheet = DispOutS1.Text;
                string newName = sheet + "_OG";
                Worksheet backUpSheet = CopyNewSheet(ThisWorkBook.Sheets[sheet], newName);
                OGSheets.Add(backUpSheet);                
            }
            return OGSheets;
        }

        private bool RenameSheetsToSave(List<Worksheet> OGSheets, string nameToAppend)
        {
            HashSet<string> PrintSheets = ((MultipleSheetsAttribute)OtherAttributeDic["SavedDupeSheet"]).GetSheetNamesHash();
            if (PrintSheets == null)
            {
                return true;
            }
            foreach (string sheet in PrintSheets)
            {
                Worksheet newSheet = ThisWorkBook.Sheets[sheet];
                string newName = newSheet.Name + " " + nameToAppend;
                // Try to rename, if unable to find name, reset OG sheets and return 
                try
                {
                    newSheet.Name = newName;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    int counter = 2;
                    bool sheetAdded = false;
                    while (counter < 101 && !sheetAdded)
                    {
                        string overlapName = newName + "(" + counter.ToString() + ")";
                        try
                        {
                            newSheet.Name = overlapName;
                            sheetAdded = true;
                        }
                        catch { }
                        counter += 1;
                    }
                    if (!sheetAdded)
                    {
                        MessageBox.Show($"Unable to create new sheet for {newName}, too many existing sheets with same name encountered");
                        ResetOGSheets(OGSheets);
                        return false;
                    }
                }
                newSheet.Move(After: ThisWorkBook.Sheets[ThisWorkBook.Sheets.Count]);
            }
            return true;
        }
        #endregion

        #endregion

        private void RunSingle_Click(object sender, EventArgs e)
        {
            #region Get Inputs
            Worksheet homeSheet = ThisApplication.ActiveSheet;
            (Range headerRange, Range InputRange, Worksheet OutputSheet) = GetSingleRunInputs();
            if (headerRange == null)
            {
                return;
            }
            
            // Convert header into list
            (List<(Range, string)> OutputHeaders, List<(Range, string)> InputHeaders) = ConvertHeaders(headerRange);
            if (OutputHeaders.Count == 0)
            {
                return;
            }
            
            // Get confirmation to continue
            DialogResult confirmation = MessageBox.Show($"Confirm to run selection for {InputRange.Rows.Count} rows?\nAny existing values in result column will be deleted.", "Confirmation", MessageBoxButtons.YesNo);
            if (confirmation == DialogResult.No) { return; }

            #endregion

            #region Rename Base Sheets
            List<Worksheet> OGSheets = RenameBaseSheets();
            if (OGSheets.Count == 0) { return; }
            #endregion

            #region Loop through all rows
            foreach (Range Row in InputRange.Rows)
            {
                int RowstoRunRowNum = Row.Row;

                #region Set and Get Values  
                bool success = GetandSetTrueValues(OutputHeaders, InputHeaders, RowstoRunRowNum, OGSheets);
                if (!success)
                {
                    return;
                }
                #endregion

                #region Rename Sheet and make new sheet if required 
                if (CheckNewSheet1.Checked)
                {
                    // Rename Sheets To Append Name
                    bool toContinue = RenameSheetsToSave(OGSheets, Row.Text);
                    if (!toContinue) { return; }
                }
                else
                {
                    string sheet = DispOutS1.Text;
                    Globals.ThisAddIn.Application.DisplayAlerts = false;
                    Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheet].Delete();
                    Globals.ThisAddIn.Application.DisplayAlerts = true;                    
                }
                // Copy new sheets for next iteration
                foreach (Worksheet OGWorksheet in OGSheets)
                {
                    string newName = OGWorksheet.Name.Substring(0, OGWorksheet.Name.Length - 3);
                    Worksheet newSheet = CopyNewSheet(OGWorksheet, newName);
                }
                #endregion
            }
            #endregion

            ResetOGSheets(OGSheets);
            homeSheet.Activate();
            MessageBox.Show("Completed run.", "Completed");
        }

        private void OverrideInputCheck_CheckedChanged(object sender, EventArgs e)
        {
            // Sets font color (grey/black) of the font of InputRange1 if override option is clicked
            if (OverrideInputCheck.Checked)
            {
                DispInputR1.ForeColor = Color.LightGray;
            }
            else
            {
                DispInputR1.ForeColor = Color.Black;
            }
        }

        #endregion


        #region Run Multiple 

        #region Helper Functions
        private (Range, Range, Range, string, string, Range, bool, bool) GetMultipleRunInputs()
        {
            // Iteration Source 
            (Range, Range, Range, string, string, Range, bool, bool) nullOutput = (null, null, null, null, null, null, false, false);
            Range iterSourceRange;
            try
            {
                iterSourceRange = GetRangeFromFullAddress(DispItSource.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for Iteration Data Source.");
                return nullOutput;
            }
            
            // Destination Columns
            Range iterDestColRange;
            try
            {
                iterDestColRange = GetRangeFromFullAddress(DispItDest.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for Destination Columns.");
                return nullOutput;
            }
            if (iterDestColRange.Rows.Count != 1)
            {
                MessageBox.Show("Invalid input for Destination Columns. Only one row allowed.");
                return nullOutput;
            }

            // Check source vs destination size 
            if (iterSourceRange.Columns.Count != iterDestColRange.Columns.Count)
            {
                MessageBox.Show("Number of columns in Iteration Data Source must match number of columns in Destination Columns");
                return nullOutput;
            }

            // Criteria Source Columns
            Range CriteriaSourceRange;
            try
            {
                CriteriaSourceRange = GetRangeFromFullAddress(DispCriteriaSource.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for Value (UR) Column.");
                return nullOutput;
            }
            if (CriteriaSourceRange.Columns.Count != 1)
            {
                MessageBox.Show("Invalid input for Value (UR) Column. Only one column allowed.");
                return nullOutput;
            }

            // LogicSymbol
            string logicSymbol;
            try
            {
                logicSymbol = DispLogicSymbol.Text;
            }
            catch
            {
                MessageBox.Show("Invalid input for logic symbol.");
                return nullOutput;
            }

            // CriteriaValue
            string CriteriaValue;
            try
            {
                CriteriaValue = DispCriteriaValue.Text;
            }
            catch
            {
                MessageBox.Show("Invalid input for Target Value (UR).");
                return nullOutput;
            }
            if (CriteriaValue == "")
            {
                MessageBox.Show("Invalid input for Target Value (UR), field cannot be empty.");
                return nullOutput;
            }
            // Try to parse UR into double
            bool isURDouble = false;
            if (double.TryParse(CriteriaValue, out double uRdouble))
            {
                isURDouble = true;
            }

            // StatusCol1
            Range statusCol;
            try
            {
                statusCol = GetRangeFromFullAddress(DispStatusCol.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for status columns.");
                return nullOutput;
            }

            // IterationMode
            bool tryAll = true; 
            try
            {
                string iterationMode = DispIterationMode.Text;
                if (iterationMode[0] == '1')
                {
                    tryAll = false;
                }
            }
            catch
            {
                MessageBox.Show("Invalid input for iteration mode.");
                return nullOutput;
            }
            return (iterSourceRange, iterDestColRange, CriteriaSourceRange,logicSymbol, CriteriaValue, statusCol, tryAll, isURDouble);

        }
        private (Range, string, string) GetOptimiseInputs()
        {
            (Range, string, string) nullOutput = (null, null, null);
            // Optimisation Column
            Range optimiseColRange;
            try
            {
                optimiseColRange = GetRangeFromFullAddress(DispOptimiseCol.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for Optimisation Source Column.");
                return nullOutput;
            }
            if (optimiseColRange.Columns.Count != 1)
            {
                MessageBox.Show("Invalid input for Optimisation Source Column. Only one column allowed.");
                return nullOutput;
            }

            // Optimisation Type
            string optimisationType;
            try
            {
                string dispText = DispOptimiseType.Text;
                if (dispText[0] == '1')
                {
                    optimisationType = "min";
                }
                else if (dispText[0] == '2')
                {
                    optimisationType = "max";
                }
                else if (dispText[0] == '3')
                {
                    optimisationType = "target";
                }
                else
                {
                    throw new Exception("Error reading optimisation type");
                }
            }
            catch
            {
                MessageBox.Show("Invalid input for optimisation type.");
                return nullOutput;
            }

            // Optimisation Target Value
            string optimisationTarget = "";
            if (optimisationType == "target")
            {
                try
                {
                    optimisationTarget = DispOptimiseTarget.Text;
                }
                catch
                {
                    MessageBox.Show("Invalid input for Optimisation Target Value.");
                    return nullOutput;
                }
                if (optimisationTarget == "")
                {
                    MessageBox.Show("Invalid input for Optimisation Target Value, field cannot be empty.");
                    return nullOutput;
                }
                else
                {
                    // Try to parse valaue 
                    bool canParse = double.TryParse(optimisationTarget, out double result);
                    if (!canParse)
                    {
                        MessageBox.Show("Invalid input for Optimisation Target Value, field must be a number.");
                        return nullOutput;
                    }
                }
            }
            
            return (optimiseColRange, optimisationType, optimisationTarget); 
        }
        private bool RangesOverlap(Range range1, Range range2, string type)
        {
            if (type == "col")
            {
                // Check if the column overlaps
                if ((range1.Column <= range2.Column && range1.Column + range1.Columns.Count - 1 >= range2.Column) ||
                    (range2.Column <= range1.Column && range2.Column + range2.Columns.Count - 1 >= range1.Column))
                {
                    return true; 
                }
                return false;
            }
            else if (type == "row")
            {
                // Check if the rows overlap
                if ((range1.Row <= range2.Row && range1.Row + range1.Rows.Count - 1 >= range2.Row) ||
                    (range2.Row <= range1.Row && range2.Row + range2.Rows.Count - 1 >= range1.Row))
                {
                    return true; 
                }
                return false;
            }
            else
            {
                throw new ArgumentException($"Input can only be 'col' or 'row', but {type} recieved");
            }
        }

        private bool OverwriteDestWithSource(Range sourceRange, Range destRange)
        {
            if (sourceRange.Rows.Count != destRange.Rows.Count || sourceRange.Columns.Count != destRange.Columns.Count)
            {
                MessageBox.Show("Source and target ranges must be of the same size.");
                return false;
            }
            
            for (int row = 1; row <= sourceRange.Rows.Count; row++)
            {
                for (int col = 1; col <= sourceRange.Columns.Count; col++)
                {
                    destRange.Cells[row, col].Value2 = sourceRange.Cells[row, col].Value2;
                }
            }
            return true;
        }

        private bool CompareValues(string currentUR, string targetUR, string logicSymbol)
        {

            #region Convert numbers to double
            bool convertCurrent = double.TryParse(currentUR, out double currentURDouble);
            bool convertTarget = double.TryParse(targetUR, out double targetURDouble);
            #endregion

            if (convertTarget && convertCurrent)
            {
                switch (logicSymbol)
                {
                    case ">": return currentURDouble > targetURDouble;
                    case "<": return currentURDouble < targetURDouble;
                    case "=": return currentURDouble == targetURDouble;
                    case ">=": return currentURDouble >= targetURDouble;
                    case "<=": return currentURDouble <= targetURDouble;
                    case "!=": return currentURDouble != targetURDouble;
                    default: throw new Exception("invalid logic");
                }
            }
            else
            {
                switch (logicSymbol)
                {
                    case "=": return currentUR == targetUR;
                    case "!=": return currentUR != targetUR;
                    default: throw new Exception($"Unable to convert {currentUR} or {targetUR} to number. Only '=' or '!=' operator is allowed.");
                }
            }
        }

        private bool IsCurrentValueBetter(string currentValueString, string bestValueString, string comparisonType, string targetValueString = "")
        {

            #region Convert numbers to double

            bool convertSuccess = false;
            convertSuccess = double.TryParse(currentValueString, out double currentValue);
            if (!convertSuccess)
            {
                throw new Exception("Unable to convert to current value from results to double");
            }

            convertSuccess = double.TryParse(bestValueString, out double bestValue);
            if (!convertSuccess)
            {
                throw new Exception("Unable to convert to best value from results to double");
            }

            #endregion

            #region Convert numbers to double
            double targetValue;
            if (comparisonType == "target")
            {
                convertSuccess = double.TryParse(targetValueString, out targetValue);
                if (!convertSuccess)
                {
                    string msg = $"Unable to convert to target value {targetValueString} from results to double";
                    MessageBox.Show(msg);
                    throw new Exception(msg);
                }
            }
            else
            {
                targetValue = 0;
            }
            #endregion

            switch (comparisonType)
            {
                case "min": return currentValue < bestValue; // Minimise
                case "max": return currentValue > bestValue; // Maximise
                case "target":
                    double currentDistance = Math.Abs(currentValue - targetValue);
                    double bestDistance = Math.Abs(bestValue - targetValue);
                    return currentDistance < bestDistance;
                default: throw new Exception($"Invalid comparison type {comparisonType}.");
            }            
        }
        #endregion

        private void MultipleRun_Click(object sender, EventArgs e)
        {
            #region Debug Settings
            bool slowIteration = checkSlowOptimisation.Checked;
            bool slowRow = checkDebugIteration.Checked;
            bool debugMode = checkDebugMode.Checked;
            int sleepDuration = 200;
            if (debugMode || slowIteration || slowRow)
            {
                ThisApplication.ScreenUpdating = true;
            }
            else
            {
                ThisApplication.ScreenUpdating = false;
            }
            
            #endregion

            #region Get Inputs
            Worksheet homeSheet = ThisApplication.ActiveSheet;
            // Universal Input
            (Range headerRange, Range InputRange, Worksheet OutputSheet) = GetSingleRunInputs();
            if (headerRange == null) { return; }

            // Multiple Run Input
            (Range iterSourceRange, Range iterDestColRange, Range CriteriaSourceRange, string logicSymbol, string CriteriaValue, Range statusCol, bool tryAll, bool isURDouble) = GetMultipleRunInputs();
            if (iterSourceRange == null) { return; }

            // Check if ranges Overlap
            if (RangesOverlap(headerRange, statusCol, "col"))
            {
                DialogResult result = MessageBox.Show("Header Row and Status Column Overlap, data might be overwritten. Continue?", "", MessageBoxButtons.YesNo);

                if (result == DialogResult.No)
                {
                    return;
                }
            }

            // Optimisation Input
            Range optimiseColRange = null;
            string optimisationType = null;
            string optimisationTarget = null;
            if (tryAll)
            {
                (optimiseColRange, optimisationType, optimisationTarget) = GetOptimiseInputs();
            }

            // Convert header into list
            (List<(Range, string)> OutputHeaders, List<(Range, string)> InputHeaders) = ConvertHeaders(headerRange);
            if (OutputHeaders.Count == 0)
            {
                return;
            }

            // Get confirmation to proceed
            DialogResult confirmation = MessageBox.Show($"Confirm to run iteration for {InputRange.Rows.Count} rows?\nAny existing values in result column will be deleted.", "Confirmation", MessageBoxButtons.YesNo);
            if (confirmation == DialogResult.No) { return; }
            #endregion

            #region Rename Base Sheets
            List<Worksheet> OGSheets = RenameBaseSheets();
            if (OGSheets.Count == 0) { return; }
            homeSheet.Activate();
            #endregion

            #region Reset Status Range
            foreach (Range currentRow in InputRange.Rows)
            {
                Range statusRange = statusCol.Worksheet.Cells[currentRow.Row, statusCol.Column];
                statusRange.Value = "Not started";
            }
            #endregion

            #region Loop de loop
            foreach (Range currentRow in InputRange.Rows)
            {
                int RowstoRunRowNum = currentRow.Row;
                bool targetReached = false;
                int iterRowNum = 1;
                double? bestTargetValue = null;
                int? bestRowNum = null;

                // Set status
                Range statusRange = statusCol.Worksheet.Cells[currentRow.Row, statusCol.Column];
                statusRange.Value = "Started Iteration";
                Range startCell;
                Range endCell;
                Range iterDestRange;
                while (!targetReached && iterRowNum <= iterSourceRange.Rows.Count)
                {
                    #region Get and set iteration values to input sheet
                    startCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column];
                    endCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column+iterDestColRange.Columns.Count-1];
                    iterDestRange = headerRange.Worksheet.Range[startCell, endCell];
                    bool success2 = OverwriteDestWithSource(iterSourceRange.Rows[iterRowNum], iterDestRange);
                    if (!success2) { return; }
                    #endregion

                    #region Set and Get non-iteration values from input sheet to destination sheet
                    bool success = GetandSetTrueValues(OutputHeaders, InputHeaders, RowstoRunRowNum, OGSheets);
                    if (!success)
                    {
                        return;
                    }
                    #endregion

                    #region Check break condition
                    Range thisURRange = CriteriaSourceRange.Worksheet.Cells[RowstoRunRowNum, CriteriaSourceRange.Column];
                    bool fulfilCondition = false;
                    try
                    {
                        fulfilCondition = CompareValues(thisURRange.Text, CriteriaValue, logicSymbol);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        ResetOGSheets(OGSheets);
                        homeSheet.Activate();
                        return;
                    }

                    #region Debug Iteration
                    if (slowIteration)
                    {
                        thisURRange.Worksheet.Activate();
                        thisURRange.Select();
                        if (debugMode)
                        {
                            DialogResult result = MessageBox.Show($"Value is: {thisURRange.Text}\nDoes it fulfil condition? {fulfilCondition}", "Pause", MessageBoxButtons.OKCancel);
                            if (result == DialogResult.Cancel)
                            {
                                ResetOGSheets(OGSheets);
                                homeSheet.Activate();
                                return;
                            }
                        }
                        else
                        {
                            System.Threading.Thread.Sleep(sleepDuration);
                        }
                    }
                    #endregion

                    if (!fulfilCondition)
                    {
                        iterRowNum += 1;
                        continue;
                    }
                    #endregion

                    #region Replace value if condition is fulfilled 
                    if (!tryAll) // If we break once condition is met
                    {
                        // Set best row and break
                        bestRowNum = iterRowNum;
                        targetReached = fulfilCondition;
                    }
                    else // If we want find optimum
                    {
                        Range thisOptimiseRange = optimiseColRange.Worksheet.Cells[RowstoRunRowNum, optimiseColRange.Column];
                        if (bestTargetValue == null) // Set best value if it doesn't exist
                        {
                            try
                            {
                                bestTargetValue = double.Parse(thisOptimiseRange.Text);
                            }
                            catch(Exception)
                            {
                                MessageBox.Show("Error encountered. Run will be terminated.\nUnable to set best target value as {thisURRange.Text} cannot be converted to number.\nCheck iteration source col\n\n","Error");
                                ResetOGSheets(OGSheets);
                                return;
                            }
                            bestRowNum = iterRowNum;
                        }
                        else // Compare to see if this value is better than best value
                        {
                            bool toReplace = IsCurrentValueBetter(thisOptimiseRange.Text, bestTargetValue.ToString(), optimisationType, optimisationTarget);
                            #region Debug Iteration 2
                            if (slowIteration)
                            {
                                if (debugMode)
                                {
                                    DialogResult result = MessageBox.Show($"Current Value: {thisOptimiseRange.Text}\nBest Value: {bestTargetValue}\nTo Replace Best Value?: {toReplace}","Debugging",MessageBoxButtons.OKCancel);
                                    if (result == DialogResult.Cancel)
                                    {
                                        ResetOGSheets(OGSheets);
                                        homeSheet.Activate();
                                        return;
                                    }
                                    
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(sleepDuration);
                                }
                            }
                            #endregion
                            if (toReplace)
                            {
                                // If it is better, overwrite existing best 
                                bestTargetValue = double.Parse(thisOptimiseRange.Text);
                                bestRowNum = iterRowNum;
                            }
                        }
                    }
                    #endregion

                    iterRowNum += 1; // for next iteration
                }

                #region Overwite cell with optimum value (or remove)
                startCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column];
                endCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column + iterDestColRange.Columns.Count - 1];
                iterDestRange = headerRange.Worksheet.Range[startCell, endCell];
                if (bestRowNum == null) // No optimum found 
                {
                    statusRange.Value = "No value found";
                    // Remove input values
                    foreach (Range cell in iterDestRange)
                    {
                        cell.ClearContents();
                    }
                }
                else
                {
                    if (tryAll)
                    {
                        // Set status
                        statusRange.Value = "Optimum value found";
                        // Set values to optimum
                        bool success2 = OverwriteDestWithSource(iterSourceRange.Rows[bestRowNum], iterDestRange);
                        if (!success2) { return; }
                        #region Set and Get Values in Destination 
                        bool success = GetandSetTrueValues(OutputHeaders, InputHeaders, RowstoRunRowNum, OGSheets);
                        if (!success)
                        {
                            return;
                        }
                        #endregion
                    }
                    else
                    {
                        statusRange.Value = "Target Reached";
                    }

                }
                #endregion

                #region Rename Sheet and make new sheet if required 
                if (CheckNewSheet1.Checked)
                {
                    // Rename Sheets To Append Name
                    bool toContinue = RenameSheetsToSave(OGSheets, currentRow.Text);
                    if (!toContinue)
                    {
                        ResetOGSheets(OGSheets);
                        return;
                    }

                    // Copy new sheets
                    foreach (Worksheet OGWorksheet in OGSheets)
                    {
                        string newName = OGWorksheet.Name.Substring(0, OGWorksheet.Name.Length - 3);
                        Worksheet backUpSheet = CopyNewSheet(OGWorksheet, newName);
                    }
                }
                #endregion
                
                #region Debug Row
                if (slowRow)
                {
                    if (debugMode)
                    {
                        DialogResult result = MessageBox.Show($"Row {RowstoRunRowNum} completed. Continue?", "Pausing", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No)
                        {
                            ResetOGSheets(OGSheets);
                            return;
                        }
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(sleepDuration);
                    }
                }
                #endregion
            }
            #endregion
            MessageBox.Show("Multiple Iteration Completed","Completed");
            ResetOGSheets(OGSheets);
            ThisApplication.ScreenUpdating = true;
            homeSheet.Activate();
        }

        private void DispIterationMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DispIterationMode.Text[0] == '1')
            {
                OptimiseGroup.BackColor = Color.Gainsboro;
            }
            else
            {
                OptimiseGroup.BackColor = Color.AliceBlue;
            }
        }
        #endregion

        #region Multiple Criteria Iteration 
        #region Insert Header
        private void InsertCriteriaTable_Click(object sender, EventArgs e)
        {
            List<string> headers = new List<string>
            {
                "Name",
                "Iteration Data Source",
                "Destination Columns",
                "Criteria Source Col",
                "Criteria logic",
                "Criteria Value",
                "Status Col",
                "Iteration Mode",
                "Optimisation Source Column",
                "Optimisation Mode",
                "Optimisation Target Value (if applicable)"
            };

            InsertHeaders(headers);
            MessageBox.Show("Completed", "Completed");
        }

        private void InsertHeaders(List<string> headers, string type = "row")
        {
            int numRow = 0;
            int numCol = 0;
            if (type == "row")
            {
                numCol = headers.Count;
                numRow = 1;
            }
            else if (type == "col")
            {
                numRow = headers.Count;
                numCol = 1;
            }
            else
            {
                MessageBox.Show($"Invalid input for type {type}", "Error");
            }

            DialogResult result = MessageBox.Show("Insert header at current selected position? " +
                $"\nIt will replace all data in {numRow} row, {numCol} columns from the start of the current selection." +
                "\nThis action cannot be undone", "Warning", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            Range selRange = ThisApplication.Selection;
            Worksheet thisSheet = selRange.Worksheet;

            int rowNum = selRange.Row;
            int colNum = selRange.Column;

            foreach (string header in headers)
            {
                thisSheet.Cells[rowNum, colNum].Value2 = header;
                if (type == "row")
                {
                    colNum++;
                }
                else if (type == "col")
                {
                    rowNum++;
                }
                
            }
        }
        #endregion

        private Dictionary<string, object> GetMultipleIterationInputs()
        {
            #region Get Data Tabel
            Range dataTable;
            try
            {
                dataTable = GetRangeFromFullAddress(DispItDataTable.Text);
            }
            catch
            {
                MessageBox.Show("Invalid input for header.");
                throw new Exception("Error getting range.");
            }
            int maxRowNum = dataTable.Rows.Count;
            #endregion

            #region Define governing dictionary 
            Dictionary<int, string> HeaderIndex = new Dictionary<int, string>
            {
                [0] = "Name",
                [1] = "iterSourceRange",
                [2] = "iterDestColRange",
                [3] = "CriteriaSourceRange",
                [4] = "logicSymbol",
                [5] = "CriteriaValue",
                [6] = "statusCol", // no
                [7] = "tryAll",
                [8] = "optimiseColRange",
                [9] = "optimisationType",
                [10] = "optimisationTarget"
            }; // Stores row num of header

            Dictionary<string, object> TableDictionary = new Dictionary<string, object>
            {
                ["Name"] = new string[maxRowNum],
                ["iterSourceRange"] = new Range[maxRowNum],
                ["iterDestColRange"] = new Range[maxRowNum],
                ["CriteriaSourceRange"] = new Range[maxRowNum],
                ["logicSymbol"] = new string[maxRowNum],
                ["CriteriaValue"] = new string[maxRowNum],
                ["statusCol"] = new Range[maxRowNum],
                ["tryAll"] = new bool[maxRowNum],
                ["isURDouble"] = new bool[maxRowNum],
                ["optimiseColRange"] = new Range[maxRowNum],
                ["optimisationType"] = new string[maxRowNum],
                ["optimisationTarget"] = new string[maxRowNum]
            }; // Stores data of header
            // List of valid symbols
            string[] validLogicSymbol = { ">", "<", "=", ">=", "<=", "!="};
            #endregion

            #region Breakdown table
            int rowNum = 0;
            foreach (Range dataRow in dataTable.Rows)
            {
                int colNum = 0;
                foreach (Range cell in dataRow.Cells)
                {
                    string headerType = HeaderIndex[colNum];
                    switch (headerType)
                    {
                        case "Name":
                            ((string[])TableDictionary[headerType])[rowNum] = cell.Value2.ToString();
                            break;

                        case "iterSourceRange":
                            Range iterSourceRange;
                            try
                            {
                                iterSourceRange = GetRangeFromFullAddress(cell.Value2.ToString());
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for Iteration Data Source.");
                                throw new Exception("Error reading Iteration Data Source.");
                            }
                            //((Range[])TableDictionary[headerType])[rowNum] = iterSourceRange;
                            Range[] thisRange = (Range[])TableDictionary[headerType];
                            thisRange[rowNum] = iterSourceRange;
                            break;

                        case "iterDestColRange":
                            Range iterDestColRange;
                            try
                            {
                                iterDestColRange = GetRangeFromFullAddress(cell.Value2.ToString());
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for Destination Columns.");
                                throw new Exception("Error reading Destination Columns.");
                            }
                            if (iterDestColRange.Rows.Count != 1)
                            {
                                MessageBox.Show("Invalid input for Destination Columns. Only one row allowed.");
                                throw new Exception("Error reading Destination Columns.");
                            }
                            ((Range[])TableDictionary[headerType])[rowNum] = iterDestColRange;
                            break;
                        case "CriteriaSourceRange":
                            Range CriteriaSourceRange;
                            try
                            {
                                CriteriaSourceRange = GetRangeFromFullAddress(cell.Value2.ToString());
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for Criteria Source Col.");
                                throw new Exception("Error reading Criteria Source Col.");
                            }
                            if (CriteriaSourceRange.Columns.Count != 1)
                            {
                                MessageBox.Show("Invalid input for Criteria Source Col. Only one column allowed.");
                                throw new Exception("Error reading Criteria Source Col.");
                            }
                            ((Range[])TableDictionary[headerType])[rowNum] = CriteriaSourceRange;
                            break;
                        case "logicSymbol":
                            string logicSymbol;
                            try
                            {
                                logicSymbol = cell.Value2.ToString();
                                if (!validLogicSymbol.Contains(logicSymbol))
                                {
                                    throw new Exception("Invalid logic symbol");
                                }
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for logic symbol.");
                                throw new Exception("Error reading logic symbol.");
                            }
                            ((string[])TableDictionary[headerType])[rowNum] = logicSymbol;
                            break;
                        case "CriteriaValue":
                            string CriteriaValue;
                            try
                            {
                                if (cell.Value2 is bool)
                                {
                                    if (cell.Value2)
                                    {
                                        CriteriaValue = "TRUE";
                                    }
                                    else
                                    {
                                        CriteriaValue = "FALSE";
                                    }
                                }
                                else
                                {
                                    CriteriaValue = cell.Value2.ToString();
                                }
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for Criteria Value.");
                                throw new Exception("Error reading Criteria Value.");
                            }
                            if (CriteriaValue == "")
                            {
                                MessageBox.Show("Invalid input for Criteria Value, field cannot be empty.");
                            }
                            ((string[])TableDictionary[headerType])[rowNum] = CriteriaValue;

                            // Try to parse value to double
                            double uRdouble;
                            bool isURDouble = false;
                            if (double.TryParse(CriteriaValue, out uRdouble))
                            {
                                isURDouble = true;
                            }
                            ((bool[])TableDictionary["isURDouble"])[rowNum] = isURDouble;

                            break;
                        case "statusCol":
                            Range statusCol;
                            try
                            {
                                statusCol = GetRangeFromFullAddress(cell.Value2.ToString());
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for status columns.");
                                throw new Exception ("Invalid input for status columns");
                            }
                            ((Range[])TableDictionary[headerType])[rowNum] = statusCol;
                            break;
                        case "tryAll":
                            bool tryAll = true;
                            try
                            {
                                string iterationMode = cell.Value2.ToString();
                                if (iterationMode[0] == '1')
                                {
                                    tryAll = false;
                                }
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for iteration mode.");
                                throw new Exception("Invalid input for iteration mode");
                            }
                            ((bool[])TableDictionary[headerType])[rowNum] = tryAll;
                            break;
                        case "optimiseColRange":
                            Range optimiseColRange;
                            try
                            {
                                optimiseColRange = GetRangeFromFullAddress(cell.Value2.ToString());
                            }
                            catch
                            {
                                MessageBox.Show("Invalid input for Optimisation Source Column.");
                                throw new Exception("Invalid input for Optimisation Source Column.");
                            }
                            if (optimiseColRange.Columns.Count != 1)
                            {
                                MessageBox.Show("Invalid input for Optimisation Source Column. Only one column allowed.");
                                throw new Exception("Invalid input for Optimisation Source Column. Only one column allowed.");
                            }
                            ((Range[])TableDictionary[headerType])[rowNum] = optimiseColRange;
                            break;
                        case "optimisationType":
                            string optimisationType;
                            try
                            {
                                string dispText = cell.Value2.ToString();
                                if (dispText[0] == '1')
                                {
                                    optimisationType = "min";
                                }
                                else if (dispText[0] == '2')
                                {
                                    optimisationType = "max";
                                }
                                else if (dispText[0] == '3')
                                {
                                    optimisationType = "target";
                                }
                                else
                                {
                                    throw new Exception("Error reading optimisation type");
                                }
                            }
                            catch
                            {
                                string msg = "Invalid input for optimisation type.";
                                MessageBox.Show(msg);
                                throw new Exception(msg);
                            }
                            ((string[])TableDictionary[headerType])[rowNum] = optimisationType;
                            break;
                        case "optimisationTarget":
                            string optimisationTarget = "";
                            string optimisationType2 = ((string[])TableDictionary["optimisationType"])[rowNum];
                            if (optimisationType2 == "target")
                            {
                                try
                                {
                                    optimisationTarget = cell.Value2.ToString();
                                }
                                catch
                                {
                                    string msg = "Invalid input for Optimisation Target Value.";
                                    MessageBox.Show(msg);
                                    throw new Exception(msg);
                                }
                                if (optimisationTarget == "")
                                {
                                    string msg = "Invalid input for Optimisation Target Value, field cannot be empty.";
                                    MessageBox.Show(msg);
                                    throw new Exception(msg);
                                }
                                else
                                {
                                    // Try to parse valaue 
                                    bool canParse = double.TryParse(optimisationTarget, out double result);
                                    if (!canParse)
                                    {
                                        string msg = "Invalid input for Optimisation Target Value, field must be a number.";
                                        MessageBox.Show(msg);
                                        throw new Exception(msg);
                                    }
                                }
                            }
                            ((string[])TableDictionary[headerType])[rowNum] = optimisationTarget;
                            break;
                        default:
                            MessageBox.Show($"No column named {headerType} found");
                            throw new Exception("No such column");
                    }
                    colNum += 1;
                }
                rowNum += 1;
            }
            #endregion
            return TableDictionary;
        }

        private void RunMultiCriteria_Click(object sender, EventArgs e)
        {
            #region Debug Settings
            bool slowIteration = checkSlowOptimisation.Checked;
            bool slowRow = checkDebugIteration.Checked;
            bool debugMode = checkDebugMode.Checked;
            //debugMode = true;
            int sleepDuration = 200;
            if (debugMode || slowIteration || slowRow)
            {
                ThisApplication.ScreenUpdating = true;
            }
            else
            {
                ThisApplication.ScreenUpdating = false;
            }
            #endregion

            #region Get Inputs
            Worksheet homeSheet = ThisApplication.ActiveSheet;
            // Universal Input
            (Range headerRange, Range InputRange, Worksheet OutputSheet) = GetSingleRunInputs();
            if (headerRange == null) { return; }

            // Multiple Run Input and optimisation input
            Dictionary<string, object> dataTable;
            try
            {
                dataTable = GetMultipleIterationInputs();
            }
            catch
            {
                MessageBox.Show("Error encountered, termination.", "Error");
                return;
            }

            // Check if ranges Overlap
            if (RangesOverlap(headerRange, ((Range[])dataTable["statusCol"])[0], "col"))
            {
                DialogResult result = MessageBox.Show("Header Row and Status Column Overlap, data might be overwritten. Continue?", "", MessageBoxButtons.YesNo);

                if (result == DialogResult.No)
                {
                    return;
                }
            }

            // Convert header into list
            (List<(Range, string)> OutputHeaders, List<(Range, string)> InputHeaders) = ConvertHeaders(headerRange);
            if (OutputHeaders.Count == 0)
            {
                return;
            }

            // Get confirmation to proceed
            DialogResult confirmation = MessageBox.Show($"Confirm to run iteration for {InputRange.Rows.Count} rows?\nAny existing values in result column will be deleted.", "Confirmation", MessageBoxButtons.YesNo);
            if (confirmation == DialogResult.No) { return; }
            #endregion

            #region Rename Base Sheets
            List<Worksheet> OGSheets = RenameBaseSheets();
            if (OGSheets.Count == 0) { return; }
            homeSheet.Activate();
            #endregion

            #region Terminate Function
            void TerminateRun(bool premature = true)
            {
                if (premature)
                {
                    MessageBox.Show("Terminated", "Terminated");
                }
                else
                {
                    MessageBox.Show("Multiple Iteration Completed", "Completed");
                }
                ResetOGSheets(OGSheets);
                ThisApplication.ScreenUpdating = true;
                homeSheet.Activate();
            }
            #endregion

            #region Reset Status Range
            foreach (Range currentRow in InputRange.Rows)
            {
                foreach (Range statusCol in (Range[])dataTable["statusCol"])
                {
                    Range statusRange = statusCol.Worksheet.Cells[currentRow.Row, statusCol.Column];
                    statusRange.Value = "Not started";
                }
            }
            #endregion

            #region Loop de loop
            // Loop through all input rows
            foreach (Range currentRow in InputRange.Rows) 
            {
                // Loop through all iteration rows (mutliple criteria rows)
                for (int iterationSetNum = 0; iterationSetNum < ((string[])dataTable["Name"]).Count(); iterationSetNum += 1)
                {
                    #region Define Standard Variables
                    int RowstoRunRowNum = currentRow.Row;
                    bool targetReached = false;
                    int iterRowNum = 1;
                    double? bestTargetValue = null;
                    int? bestRowNum = null;
                    #endregion

                    #region Define Variables in Dictionary 
                    string name = ((string[])dataTable["Name"])[iterationSetNum];
                    Range iterSourceRange = ((Range[])dataTable["iterSourceRange"])[iterationSetNum];
                    Range iterDestColRange = ((Range[])dataTable["iterDestColRange"])[iterationSetNum];
                    Range CriteriaSourceRange = ((Range[])dataTable["CriteriaSourceRange"])[iterationSetNum];
                    string logicSymbol = ((string[])dataTable["logicSymbol"])[iterationSetNum];
                    string CriteriaValue = ((string[])dataTable["CriteriaValue"])[iterationSetNum];
                    Range statusCol = ((Range[])dataTable["statusCol"])[iterationSetNum];
                    bool tryAll = ((bool[])dataTable["tryAll"])[iterationSetNum];
                    bool isURDouble = ((bool[])dataTable["isURDouble"])[iterationSetNum];

                    Range optimiseColRange = ((Range[])dataTable["optimiseColRange"])[iterationSetNum];
                    string optimisationType = ((string[])dataTable["optimisationType"])[iterationSetNum];
                    string optimisationTarget = ((string[])dataTable["optimisationTarget"])[iterationSetNum];
                    #endregion

                    #region RunMultiple
                    bool RunMultiple()
                    {
                        // Set status
                        Range statusRange = statusCol.Worksheet.Cells[currentRow.Row, statusCol.Column];
                        statusRange.Value = "Started Iteration";
                        Range startCell;
                        Range endCell;
                        Range iterDestRange;
                        while (!targetReached && iterRowNum <= iterSourceRange.Rows.Count)
                        {
                            #region Get and set iteration values to input sheet
                            startCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column];
                            endCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column + iterDestColRange.Columns.Count - 1];
                            iterDestRange = headerRange.Worksheet.Range[startCell, endCell];
                            bool success2 = OverwriteDestWithSource(iterSourceRange.Rows[iterRowNum], iterDestRange);
                            if (!success2) { return false; }

                            #endregion

                            #region Set and Get non-iteration values from input sheet to destination sheet
                            bool success = GetandSetTrueValues(OutputHeaders, InputHeaders, RowstoRunRowNum, OGSheets);
                            if (!success)
                            {
                                return false;
                            }
                            #endregion

                            #region Check break condition
                            Range thisURRange = CriteriaSourceRange.Worksheet.Cells[RowstoRunRowNum, CriteriaSourceRange.Column];
                            bool fulfilCondition = false;
                            try
                            {
                                fulfilCondition = CompareValues(thisURRange.Text, CriteriaValue, logicSymbol);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                                //ResetOGSheets(OGSheets);
                                //homeSheet.Activate();
                                return false;
                            }

                            #region Debug Iteration
                            if (slowIteration)
                            {
                                thisURRange.Worksheet.Activate();
                                thisURRange.Select();
                                if (debugMode)
                                {
                                    DialogResult result = MessageBox.Show($"Value is: {thisURRange.Text}\nDoes it fulfil condition? {fulfilCondition}", "Pause", MessageBoxButtons.OKCancel);
                                    if (result == DialogResult.Cancel)
                                    {
                                        //ResetOGSheets(OGSheets);
                                        //homeSheet.Activate();
                                        return false;
                                    }
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(sleepDuration);
                                }
                            }
                            #endregion

                            if (!fulfilCondition)
                            {
                                iterRowNum += 1;
                                continue;
                            }
                            #endregion

                            #region Replace value if condition is fulfilled 
                            if (!tryAll) // If we break once condition is met
                            {
                                // Set best row and break
                                bestRowNum = iterRowNum;
                                targetReached = fulfilCondition;
                            }
                            else // If we want find optimum
                            {
                                Range thisOptimiseRange = optimiseColRange.Worksheet.Cells[RowstoRunRowNum, optimiseColRange.Column];
                                if (bestTargetValue == null) // Set best value if it doesn't exist
                                {
                                    try
                                    {
                                        bestTargetValue = double.Parse(thisOptimiseRange.Text);
                                    }
                                    catch (Exception)
                                    {
                                        MessageBox.Show($"Error encountered. Run will be terminated.\nUnable to set best target value as {thisOptimiseRange.Text} cannot be converted to number.\nCheck iteration source col\n\n", "Error");
                                        //ResetOGSheets(OGSheets);
                                        return false;
                                    }
                                    bestRowNum = iterRowNum;
                                }
                                else // Compare to see if this value is better than best value
                                {
                                    bool toReplace = IsCurrentValueBetter(thisOptimiseRange.Text, bestTargetValue.ToString(), optimisationType, optimisationTarget);
                                    #region Debug Iteration 2
                                    if (slowIteration)
                                    {
                                        if (debugMode)
                                        {
                                            DialogResult result = MessageBox.Show($"Current Value: {thisOptimiseRange.Text}\nBest Value: {bestTargetValue}\nTo Replace Best Value?: {toReplace}", "Debugging", MessageBoxButtons.OKCancel);
                                            if (result == DialogResult.Cancel)
                                            {
                                                //ResetOGSheets(OGSheets);
                                                //homeSheet.Activate();
                                                return false;
                                            }

                                        }
                                        else
                                        {
                                            System.Threading.Thread.Sleep(sleepDuration);
                                        }
                                    }
                                    #endregion
                                    if (toReplace)
                                    {
                                        // If it is better, overwrite existing best 
                                        bestTargetValue = double.Parse(thisOptimiseRange.Text);
                                        bestRowNum = iterRowNum;
                                    }
                                }
                            }
                            #endregion

                            iterRowNum += 1; // for next iteration
                        }

                        #region Overwite cell with optimum value (or remove)
                        startCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column];
                        endCell = headerRange.Worksheet.Cells[RowstoRunRowNum, iterDestColRange.Column + iterDestColRange.Columns.Count - 1];
                        iterDestRange = headerRange.Worksheet.Range[startCell, endCell];
                        if (bestRowNum == null) // No optimum found 
                        {
                            statusRange.Value = "No value found";
                            // Remove input values
                            foreach (Range cell in iterDestRange)
                            {
                                cell.ClearContents();
                            }
                        }
                        else
                        {
                            if (tryAll)
                            {
                                // Set status
                                statusRange.Value = "Optimum value found";
                                // Set values to optimum
                                bool success2 = OverwriteDestWithSource(iterSourceRange.Rows[bestRowNum], iterDestRange);
                                if (!success2) { return false; }
                                #region Set and Get Values in Destination 
                                bool success = GetandSetTrueValues(OutputHeaders, InputHeaders, RowstoRunRowNum, OGSheets);
                                if (!success)
                                {
                                    return false;
                                }
                                #endregion
                            }
                            else
                            {
                                statusRange.Value = "Target Reached";
                            }

                        }
                        #endregion

                        #region Rename Sheet and make new sheet if required (only last iteration)
                        if (CheckNewSheet1.Checked && 
                            (iterationSetNum == ((string[])dataTable["Name"]).Count()-1))
                        {
                            // Rename Sheets To Append Name
                            bool toContinue = RenameSheetsToSave(OGSheets, currentRow.Text);
                            if (!toContinue)
                            {
                                //ResetOGSheets(OGSheets);
                                return false;
                            }

                            // Copy new sheets
                            foreach (Worksheet OGWorksheet in OGSheets)
                            {
                                string newName = OGWorksheet.Name.Substring(0, OGWorksheet.Name.Length - 3);
                                Worksheet backUpSheet = CopyNewSheet(OGWorksheet, newName);
                            }
                        }
                        #endregion

                        #region Debug Row
                        if (slowRow)
                        {
                            if (debugMode)
                            {
                                DialogResult result = MessageBox.Show($"Row {RowstoRunRowNum} completed. Continue?", "Pausing", MessageBoxButtons.YesNo);
                                if (result == DialogResult.No)
                                {
                                    //ResetOGSheets(OGSheets);
                                    return false;
                                }
                            }
                            else
                            {
                                System.Threading.Thread.Sleep(sleepDuration);
                            }
                        }
                        #endregion
                        return true;
                    }

                    bool isSuccess = RunMultiple();
                    if (!isSuccess)
                    {
                        TerminateRun();
                        return;
                    }
                    #endregion
                }
            }
            #endregion
            TerminateRun(false);
            //MessageBox.Show("Multiple Iteration Completed", "Completed");
            //ResetOGSheets(OGSheets);
            //ThisApplication.ScreenUpdating = true;
            //homeSheet.Activate();
        }

        #endregion

        #region Sheet Management
        private void duplicateSheets_Click(object sender, EventArgs e)
        {
            #region Check Input Sheet Is Set
            Worksheet copySheet;
            try
            {
                copySheet = ((SheetTextBox)RangeAttributeDic["CopySheet"]).getSheet();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }

            #endregion

            #region Read Table Info
            int numCopied = 0;
            try
            {
                ThisApplication.ScreenUpdating = false;
                Range selRange = ThisApplication.ActiveWindow.RangeSelection;
                foreach (Range cell in selRange)
                {
                    if (cell.Value2 != null && cell.Value2 != "")
                    {
                        Worksheet newSheet = CopyNewSheet(copySheet, cell.Value2);
                        newSheet.Move(After: ThisWorkBook.Sheets[ThisWorkBook.Sheets.Count]);
                        numCopied++;
                    }
                }
            }
            catch 
            {
            }
            finally
            {
                ThisApplication.ScreenUpdating = true;
            }
            
            #endregion

            #region Check if copy is done
            if (numCopied == 0)
            {
                MessageBox.Show("Failed to copy", "Error");
            }
            else
            {
                MessageBox.Show($"Copied {numCopied} sheets", "Completed");
            }
            #endregion
        }

        private void getSheetNames_Click(object sender, EventArgs e)
        {
            #region Get Sheet Info
            List<string> names = new List<string>();
            foreach (Worksheet worksheet in ThisWorkBook.Worksheets)
            {
                names.Add(worksheet.Name);
            }
            #endregion

            #region Confirmation
            try
            {
                WriteToExcelSelectionAsRow(0, 0, true, names.ToArray());
                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            #endregion
        }

        private void setSheetNames_Click(object sender, EventArgs e)
        {
            Workbook workbook = ThisApplication.ActiveWorkbook;
            Range selectedRange = ThisApplication.ActiveWindow.RangeSelection;
            
            #region Checks
            try
            {
                CheckRangeSize(selectedRange, 0, 2);
                List<string> ogSheetNamesL = GetContentsAsStringList(selectedRange.Columns[1].Cells, true);
                List<string> newSheetNamesL = GetContentsAsStringList(selectedRange.Columns[2].Cells, true);
                CheckIfSheetsExist(workbook, ogSheetNamesL, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
            #endregion
            string[] ogSheetNames = GetContentsAsStringArray(selectedRange.Columns[1].Cells, false);
            string[] newSheetNames = GetContentsAsStringArray(selectedRange.Columns[2].Cells, false);
            List<string> failedToRename = new List<string>();
            string[] finalSheetNames = new string[ogSheetNames.Count()];
            for (int i = 0; i < ogSheetNames.Length; i++)
            {
                string ogName = ogSheetNames[i];
                string newName = newSheetNames[i];

                try
                {
                    Worksheet sheet = ThisWorkBook.Sheets[ogName];
                    sheet.Name = newName;
                    finalSheetNames[i] = newSheetNames[i];
                }
                catch (Exception ex)
                {
                    failedToRename.Add($"{ogName}: {ex.Message}");
                    finalSheetNames[i] = ogSheetNames[i];
                }
            }

            if (failedToRename.Count > 0)
            {
                MessageBox.Show($"Failed to rename the following sheets:\n{ConvertToString(failedToRename)}", "Error");
                DialogResult res = MessageBox.Show($"Print final sheet names to adjacent cells?", "Confirmation", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                {
                    WriteToExcelSelectionAsRow(0, 2, false, finalSheetNames);
                }
            }
            else
            {
                MessageBox.Show($"Completed", "Completed");
            }
        }

        private void reorderSheets_Click(object sender, EventArgs e)
        {
            Workbook workbook = ThisApplication.ActiveWorkbook;
            Range selectedRange = ThisApplication.ActiveWindow.RangeSelection;
            Worksheet currentSheet = ThisApplication.ActiveSheet;
            List<string> sheetNames;

            #region Checks
            try
            {
                CheckRangeSize(selectedRange, 0, 1);
                sheetNames = GetContentsAsStringList(selectedRange.Columns[1].Cells, true);
                CheckIfSheetsExist(workbook, sheetNames, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
            #endregion

            List<string> failed = new List<string>();
            for (int i = 1; i < sheetNames.Count; i++)
            {
                string prevSheetName = sheetNames[i-1];
                string sheetName = sheetNames[i];

                try
                {
                    Worksheet thisSheet = ThisWorkBook.Sheets[sheetName];
                    Worksheet prevSheet = ThisWorkBook.Sheets[prevSheetName];
                    thisSheet.Move(After: prevSheet);
                }
                catch (Exception ex)
                {
                    failed.Add($"{sheetName}: {ex.Message}");
                }
            }

            currentSheet.Activate();

            if (failed.Count > 0)
            {
                MessageBox.Show($"Failed to move the following sheets:\n{ConvertToString(failed)}", "Error");
            }
            else
            {
                MessageBox.Show($"Completed", "Completed");
            }
        }
        #endregion

        #region Misc
        private void convertToValue_Click(object sender, EventArgs e)
        {
            Range selRange = ThisApplication.ActiveWindow.RangeSelection;
            foreach (Range cell in selRange)
            {
                cell.Value = cell.Value;
            }
            MessageBox.Show("Completed", "Completed");
        }

        private void shiftValuesDown_Click(object sender, EventArgs e)
        {
            Range selRange = ThisApplication.ActiveWindow.RangeSelection;
            (int startRow, int endRow, int startCol, int endCol) = CommonUtilities.GetRangeDetails(selRange);

            for (int rowNum = endRow; rowNum >= startRow; rowNum--)
            {
                for (int colNum = startCol; colNum <= endCol; colNum++)
                {
                    Range cell = selRange.Worksheet.Cells[rowNum, colNum];
                    Range newCell = cell.Offset[1, 0];
                    newCell.Value2 = cell.Value2;

                    if (checkClearNewCell.Checked)
                    {
                        if (rowNum == startRow)
                        {
                            cell.ClearContents();
                        }
                    }
                }
            }

            MessageBox.Show("Completed", "Completed");
        }
        #endregion

        #region Single Cell Interation
        private void increaseVal_Click(object sender, EventArgs e)
        {
            runSingleIter(true);
        }

        private void decreaseVal_Click(object sender, EventArgs e)
        {
            runSingleIter(false);
        }

        private void runSingleIter(bool increaseVal)
        {
            #region Get Inputs
            Range selRange = ThisApplication.ActiveWindow.RangeSelection;
            if (selRange.Rows.Count > 1 || selRange.Columns.Count > 1)
            {
                MessageBox.Show("Selected range can only contain one cells", "Error");
                return;
            }

            #region Create Iteration Object
            TargetCriteria criteria = new TargetCriteria((RangeTextBox)RangeAttributeDic["CriteriaSource2"], (ComboBoxAttribute)OtherAttributeDic["LogicSymbol2"], RangeAttributeDic["CriteriaValue2"]);

            if (criteria.CriteriaMet())
            {
                MessageBox.Show("Criteria already fulfilled", "Error");
            }
            #endregion

            double increment;
            double maxIter;
            try
            {
                increment = RangeAttributeDic["Increment"].GetDoubleFromTextBox();
                maxIter = RangeAttributeDic["LoopNum"].GetIntFromTextBox();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Error");
                return;
            }
            #endregion

            for (int iterNum = 0; iterNum <maxIter; iterNum++)
            {
                selRange.Value2 = selRange.Value2 + increment;
                if (criteria.CriteriaMet())
                {
                    MessageBox.Show("Value found", "Completed");
                    return;
                }
            }
            MessageBox.Show($"Max number of loops reached\nValue not found", "Error");
        }

        private void createIterObject()
        {
            TargetCriteria criteria = new TargetCriteria((RangeTextBox)RangeAttributeDic["CriteriaSource2"], (ComboBoxAttribute)OtherAttributeDic["LogicSymbol2"], RangeAttributeDic["CriteriaValue2"]);
            criteria.CheckInputs();
            bool pass = criteria.CriteriaMet();

        }
        #endregion

        #region Testing Grounds
        public TabPage GetPageTaskPane(int tabNum)
        {
            TabControl.TabPageCollection MyTabPages = ExcelTabControl.TabPages;
            TabPage ThisTabPage = MyTabPages[tabNum];
            MessageBox.Show(ThisTabPage.Name);
            return ThisTabPage;
        }

        #endregion
    }
}


