using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using ExcelAddIn2.Excel_Pane_Folder;
using Button = System.Windows.Forms.Button;
using TextBox = System.Windows.Forms.TextBox;
using System.IO;

namespace ExcelAddIn2
{
    public class AttributeTextBox
    {
        #region Initialisation
        public string attName;
        public System.Windows.Forms.TextBox textBox;
        public string type = "string";
        public string defaultValue;
        public AttributeTextBox(string attName, TextBox textBox, bool isBasicValue = false)
        {
            this.attName = attName;
            this.textBox = textBox;

            if (isBasicValue)
            {
                SubscribeToTextBoxEvents();
            }

            RefreshTextBox();
        }

        protected void SubscribeToTextBoxEvents()
        {
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        protected void textBox_LostFocus(object sender, EventArgs e)
        {
            PerformVerification();
            SetValueFromTextBox();
        }

        protected void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
        }
        #endregion

        #region Additional Verification
        
        private void PerformVerification()
        {
            if (textBox.Text == "")
            {
                return;
            }

            switch (type)
            {
                case "string":
                    break;
                case "int":
                    CheckIsInt();
                    break;
                case "double":
                    CheckIsDouble();
                    break;
                case "filename":
                    CheckIsFileName();
                    break;
                case "partial filepath":
                    CheckIsPartialFilePath();
                    break;
                default:
                    throw new Exception("Verification type not found");
            }
        }

        private bool CheckIsInt()
        {
            try
            {
                Convert.ToInt32(textBox.Text);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Value '{textBox.Text}' for {attName} cannot be converted to integer\n\n"+ex.Message, "Error");
                RefreshTextBox();
                return false;
            }
        }

        private bool CheckIsDouble()
        {
            bool success = double.TryParse(textBox.Text, out double value);
            if (!success)
            {
                MessageBox.Show($"Value '{textBox.Text}' for {attName} cannot be converted to number (double)", "Error");
                RefreshTextBox();
                return false;
            }
            return true;   
        }

        private bool CheckIsFileName()
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidFileNameChars());
            bool success = true;
            string invalidCharsInString = "";
            foreach(char c in textBox.Text)
            {
                if (invalidChars.Contains(c))
                {
                    invalidCharsInString += c + " ";
                    success = false;
                }
            }

            if (!success)
            {
                MessageBox.Show($"'{textBox.Text}' contains invalid characters: \n"+ invalidCharsInString, "Error");
                RefreshTextBox();
                return false;
            }
            return true;
        }

        private bool CheckIsPartialFilePath()
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidPathChars());
            bool success = true;
            foreach (char c in textBox.Text)
            {
                if (invalidChars.Contains(c))
                {
                    success = false;
                    break;
                }
            }

            if (!success)
            {
                string invalidString = "";
                foreach (char character in invalidChars)
                {
                    invalidString += character;
                }
                MessageBox.Show($"'{textBox.Text}' contains invalid characters: \n{invalidString}", "Error");
                RefreshTextBox();
                return false;
            }
            return true;
        }

        #endregion

        #region Default Value
        public void SetDefaultValue(string value)
        {
            defaultValue = value;
            RefreshTextBox();
        }
        #endregion

        #region Get and Set Values for Properties
        public (bool, DocumentProperty) GetValueFromProp()
        {
            // This always gets value from what is saved in the document. 
            // First output is boolean stating if property has been set or not 
            // Second output returns the DocumentProperty type
            DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
            foreach (DocumentProperty prop in AllCustProps)
            {
                if (prop.Name == attName)
                {
                    return (true, prop);
                }
            }
            return (false, null);
        }

        public bool SetValue(object AttValue, bool showMsg = false)
        {
            try
            {
                DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
                (bool exist, DocumentProperty prop) = GetValueFromProp();
                if (exist)
                {
                    if (prop.Value.ToString() == AttValue.ToString())
                    {
                        return true;
                    }
                    prop.Value = AttValue;
                    if (showMsg)
                    {
                        MessageBox.Show(attName + " updated to " + AttValue.ToString());
                    }
                }
                else
                {
                    AllCustProps.Add(attName, false, MsoDocProperties.msoPropertyTypeString, AttValue.ToString());
                    if (showMsg)
                    {
                        MessageBox.Show(attName + " added as " + AttValue.ToString());
                    }
                }
                RefreshTextBox();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding value {AttValue} for {attName}: \n\n{ex.Message}", "Error");
                RefreshTextBox();
                return false;
            }
        }

        public void RefreshTextBox()
        {
            (bool gotValue, DocumentProperty ThisProp) = GetValueFromProp();
            if (gotValue)
            {
                string DisplayValue = ThisProp.Value.ToString();
                textBox.Text = DisplayValue;
            }
            else
            {
                if (defaultValue != null)
                {
                    textBox.Text = defaultValue;
                    //SetValueFromTextBox();
                }
                else
                {
                    textBox.Clear();
                }
            }
        }

        protected virtual void SetValueFromTextBox(bool showMsg = false)
        {
            SetValue(textBox.Text);
        }
        #endregion

        #region Get Values for TextBox
        public double GetDoubleFromTextBox()
        {
            bool check = double.TryParse(textBox.Text, out double doubleValue);
            if (!check)
            {
                throw new Exception($"Unable to parse {textBox.Text} into number for {attName}");
            }
            else
            {
                return doubleValue;
            }
        }

        public float GetFloatFromTextBox()
        {
            bool check = float.TryParse(textBox.Text, out float floatValue);
            if (!check)
            {
                throw new Exception($"Unable to parse {textBox.Text} into number for {attName}");
            }
            else
            {
                return floatValue;
            }
        }

        public int GetIntFromTextBox()
        {
            bool check = int.TryParse(textBox.Text, out int value);
            if (!check)
            {
                throw new Exception($"Unable to parse {textBox.Text} into integer for {attName}");
            }
            else
            {
                return value;
            }
        }
        #endregion

        #region Import Export
        public void ResetValue()
        {
            (bool gotValue, DocumentProperty thisProp) = GetValueFromProp();
            if (gotValue)
            {
                thisProp.Delete();
            }
            RefreshTextBox();
            if (defaultValue != null)
            {
                textBox.Text = defaultValue;
            }
        }

        public virtual bool ImportValue(string value)
        {
            try
            {
                textBox.Text = value;
                SetValueFromTextBox(false);
                return true;
            }
            catch
            { 
                return false;
            }
        }
        #endregion
    }

    public class RangeTextBox : AttributeTextBox
    {
        #region Initialising
        public TextBox rangeTextBox;
        public Button setRangeButton;
        string rangeType;
        bool withSheet;
        public RangeTextBox(string attName, TextBox textBox, Button button, string type = "range", bool withSheet = true) : base(attName, textBox)
        {
            setRangeButton = button;
            rangeTextBox = textBox;
            rangeType = type;
            this.withSheet = withSheet;
            SubscribeToRangeTextBoxEvents();
        }

        private void SubscribeToRangeTextBoxEvents()
        {
            setRangeButton.Click += new EventHandler(setButton_click);
            rangeTextBox.LostFocus += new EventHandler(rangeTextBox_LostFocus);
            rangeTextBox.KeyDown += new KeyEventHandler(rangeTextBox_KeyDown);
        }

        private void setButton_click(object sender, EventArgs e)
        {
            //switch (rangeType)
            //{
            //    case "range":
            //        SetRange(true);
            //        break;
            //    case "row":
            //        SetOneRow(true);
            //        break;
            //    case "column":
            //        SetOneCol(true);
            //        break;
            //    default:
            //        throw new Exception($"Type of rangebox for {attName} is not valid");
            //}
            SetRange2(true);
        }

        private void rangeTextBox_LostFocus(object sender, EventArgs e)
        {
            SetRangeFromTextBox2(true);
        }
        private void rangeTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
            //if (e.KeyCode == Keys.Enter)
            //{
            //    SetRangeFromTextBox2(true);
            //    SelectRange(false);
            //}
        }
        #endregion

        #region Set Range 
        public void SelectRange(bool forceSelect = true)
        {
            Range excelRange = null;
            try
            {
                string[] parts = rangeTextBox.Text.Split('!');
                if (parts.Length == 2)
                {
                    string sheetName = parts[0];
                    string range = parts[1];
                    // Check if valid worksheet and range
                    Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    excelRange = worksheet.Range[range];
                    if (forceSelect)
                    {
                        worksheet.Activate();
                    }
                    excelRange.Select();
                }
            }
            catch
            {
                if (forceSelect)
                {
                    throw new Exception($"Unable to select cell {rangeTextBox.Text}");
                }
            }
        }

        public void SetRange2(bool showMsg = false, bool forceSelect = false)
        {
            try
            {
                string PrevValue = "";
                (bool ret, DocumentProperty prop) = GetValueFromProp();
                if (ret)
                {
                    PrevValue = prop.Value.ToString();
                    SelectRange(forceSelect);
                }

                Range selectedRange;
                switch (rangeType)
                {
                    case "range":
                        {
                            var userInput = Globals.ThisAddIn.Application.InputBox("Select data range", "Select data", PrevValue, Type: 8);
                            if (userInput is bool)
                            {
                                RefreshTextBox();
                                return;
                            }
                            selectedRange = userInput;
                            break;
                        }
                    case "row":
                        {
                            var userInput = Globals.ThisAddIn.Application.InputBox("Only first row of data will be kept.", "Select range in one row", PrevValue, Type: 8);
                            if (userInput is bool)
                            {
                                RefreshTextBox();
                                return;
                            }
                            selectedRange = userInput.Rows[1];
                            break;
                        }
                    case "column":
                        {
                            var userInput = Globals.ThisAddIn.Application.InputBox("Only first column of data will be kept.", "Select range in one column", PrevValue, Type: 8);
                            if (userInput is bool)
                            {
                                RefreshTextBox();
                                return;
                            }
                            selectedRange = userInput.Columns[1];
                            break;
                        }
                    case "cell":
                        {
                            var userInput = Globals.ThisAddIn.Application.InputBox("Only first cell will be kept.", "Select a cell", PrevValue, Type: 8);
                            if (userInput is bool)
                            {
                                RefreshTextBox();
                                return;
                            }
                            selectedRange = userInput.Cells[1,1];
                            break;
                        }
                    default:
                        throw new Exception($"Type of rangebox for {attName} is not valid");
                }

                string address = selectedRange.Address[false,false];
                if (withSheet)
                {
                    address = selectedRange.Worksheet.Name + "!" + address;
                }

                SetValue(address, showMsg);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void SetRangeFromTextBox2(bool showMsg = false)
        {
            // Check existing property, if it is the same as current, return
            try
            {
                (bool ret, DocumentProperty prop) = GetValueFromProp();
                if (prop.Value == rangeTextBox.Text)
                {
                    return;
                }
            }
            catch (Exception) { }

            // If input is empty, delete value
            if (rangeTextBox.Text == "")
            {
                ResetValue();
                return;
            }

            // If input is different, try to update value stored
            try
            {
                string fullAddress = rangeTextBox.Text;
                if (!withSheet)
                {
                    string sheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                    fullAddress = sheetName + "!" + fullAddress;
                }

                string[] parts = fullAddress.Split('!');
                if (parts.Length == 2)
                {
                    string sheetName = parts[0];
                    string range = parts[1];
                    // Check if valid worksheet and range
                    Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    Range selectedRange;
                    switch (rangeType)
                    {
                        case "range":
                            selectedRange = worksheet.Range[range];
                            break;
                        case "row":
                            selectedRange = worksheet.Range[range].Rows[1];
                            break;
                        case "column":
                            selectedRange = worksheet.Range[range].Columns[1];
                            break;
                        case "cell":
                            selectedRange = worksheet.Range[range].Cells[1,1];
                            break;
                        default:
                            throw new Exception($"Type of rangebox for {attName} is not valid");
                    }

                    string address = selectedRange.Address[false,false];
                    if (withSheet)
                    {
                        address = selectedRange.Worksheet.Name + "!" + address;
                    }

                    // Set value if valid
                    bool ret = SetValue(address, showMsg);
                    RefreshTextBox();
                }
                else
                {
                    throw new Exception("Wrong Fromat");
                }
            }
            catch (Exception)
            {
                RefreshTextBox();
                MessageBox.Show($"Error: Invalid input format for {rangeTextBox.Text}, value = {rangeTextBox.Text}");
            }
        }

        #region Archival
        //public override void SetValueFromTextBox(bool showMsg = false)
        //{
        //    switch (rangeType)
        //    {
        //        case "range":
        //            SetRangeFromTextBox(showMsg);
        //            break;
        //        case "row":
        //            SetOneRowFromTextBox(showMsg);
        //            break;
        //        case "column":
        //            SetOneColFromTextBox(showMsg);
        //            break;
        //        default:
        //            throw new Exception($"Type of rangebox for {attName} is not valid");
        //    }
        //}

        //public void SetRange(bool showMsg = false)
        //{
        //    try
        //    {
        //        string PrevValue = "";
        //        (bool ret, DocumentProperty prop) = GetValue();
        //        if (ret)
        //        {
        //            PrevValue = prop.Value.ToString();
        //        }

        //        Range selectedRange = Globals.ThisAddIn.Application.InputBox("Select row(s) of data", "Select row(s)", PrevValue, Type: 8);
        //        string fullAddress = selectedRange.Worksheet.Name + "!" + selectedRange.Address;
        //        SetValue(fullAddress, showMsg);
        //    }
        //    catch (Exception)
        //    {
        //        //MessageBox.Show(ex.ToString());
        //    }
        //}
        //public void SetRangeFromTextBox(bool showMsg = false)
        //{
        //    try
        //    {
        //        (bool ret, DocumentProperty prop) = GetValue();
        //        if (prop.Value == rangeTextBox.Text)
        //        {
        //            //MessageBox.Show("Skipped updating");
        //            return;
        //        }
        //    }
        //    catch (Exception) { }

        //    // If input is different, try to update value stored
        //    try
        //    {
        //        string[] parts = rangeTextBox.Text.Split('!');
        //        if (parts.Length == 2)
        //        {
        //            string sheetName = parts[0];
        //            string range = parts[1];
        //            // Check if valid worksheet and range
        //            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
        //            Range excelRange = worksheet.Range[range];

        //            // Set value if valid
        //            bool ret = SetValue(rangeTextBox.Text, showMsg);
        //        }
        //        else
        //        {
        //            throw new Exception("Wrong Fromat");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        RefreshTextBox();
        //        MessageBox.Show($"Error: Invalid input format for {rangeTextBox.Text}, value = {rangeTextBox.Text}");
        //    }
        //}

        //public void SetOneRow(bool showMsg = false)
        //{
        //    try
        //    {
        //        string PrevValue = "";
        //        (bool ret, DocumentProperty prop) = GetValue();
        //        if (ret)
        //        {
        //            PrevValue = prop.Value.ToString();
        //        }

        //        Range selectedRange = Globals.ThisAddIn.Application.InputBox("Only first row of data will be kept.", "Select range in one row", PrevValue, Type: 8).Rows[1];
        //        string fullAddress = selectedRange.Worksheet.Name + "!" + selectedRange.Address;
        //        SetValue(fullAddress, showMsg);
        //    }
        //    catch (Exception)
        //    {
        //        //MessageBox.Show(ex.ToString());
        //    }
        //}
        //public void SetOneRowFromTextBox(bool showMsg = false)
        //{
        //    try
        //    {
        //        (bool ret, DocumentProperty prop) = GetValue();
        //        if (prop.Value == rangeTextBox.Text)
        //        {
        //            //MessageBox.Show("Skipped updating");
        //            return;
        //        }
        //    }
        //    catch (Exception) { }

        //    // If input is different, try to update value stored
        //    try
        //    {
        //        string[] parts = rangeTextBox.Text.Split('!');
        //        if (parts.Length == 2)
        //        {
        //            string sheetName = parts[0];
        //            string range = parts[1];
        //            // Check if valid worksheet and range
        //            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
        //            Range selectedRange = worksheet.Range[range].Rows[1];

        //            string fullAddress = selectedRange.Worksheet.Name + "!" + selectedRange.Address;
        //            // Set value if valid
        //            SetValue(fullAddress, showMsg);
        //            RefreshTextBox();
        //        }
        //        else
        //        {
        //            throw new Exception("Wrong Fromat");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        RefreshTextBox();
        //        MessageBox.Show($"Error: Invalid input format for {rangeTextBox.Text}, value = {rangeTextBox.Text}");
        //    }
        //}

        //public void SetOneCol(bool showMsg = false)
        //{
        //    try
        //    {
        //        string PrevValue = "";
        //        (bool ret, DocumentProperty prop) = GetValue();
        //        if (ret)
        //        {
        //            PrevValue = prop.Value.ToString();
        //        }

        //        Range selectedRange = Globals.ThisAddIn.Application.InputBox("Only first column of data will be kept.", "Select range in one column", PrevValue, Type: 8).Columns[1];
        //        string fullAddress = selectedRange.Worksheet.Name + "!" + selectedRange.Address;
        //        SetValue(fullAddress, showMsg);
        //    }
        //    catch (Exception)
        //    {
        //        //MessageBox.Show(ex.ToString());
        //    }
        //}
        //public void SetOneColFromTextBox(bool showMsg = false)
        //{
        //    try
        //    {
        //        (bool ret, DocumentProperty prop) = GetValue();
        //        if (prop.Value == rangeTextBox.Text)
        //        {
        //            //MessageBox.Show("Skipped updating");
        //            return;
        //        }
        //    }
        //    catch (Exception) { }

        //    // If input is different, try to update value stored
        //    try
        //    {
        //        string[] parts = rangeTextBox.Text.Split('!');
        //        if (parts.Length == 2)
        //        {
        //            string sheetName = parts[0];
        //            string range = parts[1];
        //            // Check if valid worksheet and range
        //            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
        //            Range selectedRange = worksheet.Range[range].Columns[1];

        //            string fullAddress = selectedRange.Worksheet.Name + "!" + selectedRange.Address;
        //            SetValue(fullAddress);

        //            // Set value if valid
        //            SetValue(fullAddress, showMsg);
        //            RefreshTextBox();
        //        }
        //        else
        //        {
        //            throw new Exception($"Wrong Fromat for {rangeTextBox.Text}");
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        RefreshTextBox();
        //        MessageBox.Show($"Error: Invalid input format for {rangeTextBox.Text}, value = {rangeTextBox.Text}");
        //    }
        //}
        #endregion
        public new bool ImportValue(string value)
        {
            try
            {
                textBox.Text = value;
                SetValueFromTextBox();
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Get Range
        public Range GetRangeFromFullAddress()
        {
            string fullAddress = rangeTextBox.Text;
            if (fullAddress == "")
            {
                throw new Exception($"Empty input for {attName}");
            }

            if (!withSheet)
            {
                string sheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                fullAddress = sheetName + "!" + fullAddress;
            }
            var parts = fullAddress.Split('!');
            if (parts.Length != 2)
            {
                throw new ArgumentException($"Invalid address format for address: {fullAddress}. Expected format: SheetName!CellAddress");
            }

            try
            {
                //ThisWorkBook.Sheets[parts[0]].Activate();
                Worksheet ThisWorksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[parts[0]];
                Range returnRange = ThisWorksheet.Range[parts[1]];
                return returnRange;
            }
            catch
            {
                MessageBox.Show($"Error Returning Range at ${fullAddress}");
                return null;
            }
        }

        public Range GetRangeForSpecificSheet(string sheetName)
        {
            #region Input Checks
            if (withSheet)
            {
                throw new Exception($"Unable to get range for {attName} with specific sheet, provided input has sheet name");
            }

            string cellAddress = rangeTextBox.Text;
            if (cellAddress == "")
            {
                throw new Exception($"Empty input for {attName}");
            }
            #endregion

            //cellAddress = sheetName + "!" + cellAddress;
            try
            {
                Worksheet ThisWorksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                Range returnRange = ThisWorksheet.Range[cellAddress];
                return returnRange;
            }
            catch
            {
                throw new Exception($"Error Returning Range at {sheetName}${cellAddress} for {attName}");
            }
        }

        public string GetFirstCellContent()
        {
            Range referencedCell = GetRangeFromFullAddress();
            return referencedCell.Cells[1, 1].Value2.ToString();
        }
        
        public Range[] GetAreaRange(string sheetName = null)
        {
            Range allRange;
            if (sheetName == null)
            {
                allRange = GetRangeFromFullAddress();
            }
            else
            {
                allRange = GetRangeForSpecificSheet(sheetName);
            }

            Range[] rangeArray = new Range[allRange.Areas.Count];
            int i = 0;
            foreach (Range range in allRange.Areas)
            {
                rangeArray[i] = range;
                i++;
            }
            return rangeArray;
        }
        #endregion

    }

    public class SheetTextBox : AttributeTextBox
    {
        System.Windows.Forms.Button setSheetButton;
        System.Windows.Forms.TextBox sheetTextBox;
        public SheetTextBox(string AttName, System.Windows.Forms.TextBox TextBox, System.Windows.Forms.Button button) : base(AttName, TextBox)
        {
            setSheetButton = button;
            sheetTextBox = TextBox;
            SubscribeToSheetTextBoxEvents();

        }
        private void SubscribeToSheetTextBoxEvents()
        {
            setSheetButton.Click += new EventHandler(setSheetButton_Click);
            sheetTextBox.LostFocus += new EventHandler(sheetTextBox_LostFocus);
            sheetTextBox.KeyDown += new KeyEventHandler(sheetTextBox_KeyDown);
        }
        private void setSheetButton_Click(object sender, EventArgs e)
        {
            SetSheet();
        }
        private void sheetTextBox_LostFocus(object sender, EventArgs e)
        {
            SetFromTextBox();
        }
        private void sheetTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SetFromTextBox();
            }
        }

        public void SetSheet()
        {
            try
            {
                string PrevValue = "";
                (bool ret, DocumentProperty prop) = GetValueFromProp();
                if (ret)
                {
                    PrevValue = prop.Value.ToString();
                }

                Range selectedRange = Globals.ThisAddIn.Application.InputBox("Select any cell in destination sheet", "Select sheet", PrevValue, Type: 8);
                string SheetName = selectedRange.Worksheet.Name;
                SetValue(SheetName);
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        public void SetFromTextBox()
        {
            try
            {
                (bool ret, DocumentProperty prop) = GetValueFromProp();
                if (prop.Value == textBox.Text)
                {
                    return;
                }
            }
            catch (Exception) { }

            // If input is different, try to update value stored
            try
            {
                Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[textBox.Text];

                string[] parts = textBox.Text.Split('!');
                bool ret = SetValue(textBox.Text);
            }
            catch (Exception)
            {
                RefreshTextBox();
                MessageBox.Show("Error: Sheet not found.");
            }
        }

        public Worksheet getSheet()
        {
            if (textBox.Text == "")
            {
                throw new Exception($"Worksheet name not set");
            }
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == textBox.Text)
                {
                    return sheet; // Return the worksheet if the name matches
                }
            }
            throw new Exception($"Worksheet {textBox.Text} not found");
        }
    }

    public class DirectoryTextBox: AttributeTextBox
    {
        #region Initialisation
        public Button setButton;
        public DirectoryTextBox(string attName, TextBox textBox, Button button) : base(attName, textBox, false)
        {
            setButton = button;
            SubscribeToDirectoryTextBoxEvents();
        }

        private void SubscribeToDirectoryTextBoxEvents()
        {
            setButton.Click += new EventHandler(setButton_Click);
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        #endregion

        #region Add Open Folder Button
        public Button openButton;
        public void AddOpenButton(Button openButton)
        {
            this.openButton = openButton;
            openButton.Click += new EventHandler(openButton_Click);
        }
        #endregion

        #region Basic Operations
        private void setButton_Click(object sender, EventArgs e)
        {
            CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
            if (customFolderBrowser.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            SetValue(customFolderBrowser.GetFolderPath());
        }

        private void openButton_Click(object sender, EventArgs e)
        {
            // Check if button has been added
            if (openButton == null)
            {
                throw new Exception("Open Button not added to this directory");
            }

            // Get Path
            string folderPath = textBox.Text;

            //Check if path exist
            if (folderPath == "")
            {
                MessageBox.Show("No path provided", "Error");
                return;
            }

            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show($"Directory provided is invalid: \n\n{folderPath}?", "Error");
                return;
            }
            else
            {
                System.Diagnostics.Process.Start(folderPath);
            }
        }

        protected new void textBox_LostFocus(object sender, EventArgs e)
        {
            SetValueFromTextBox();
        }

        protected new void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
        }

        protected new void SetValueFromTextBox(bool showMsg = false)
        {
            if (textBox.Text == "")
            {
                ResetValue();
                return;
            }

            if (!Path.IsPathRooted(textBox.Text) || !Directory.Exists(textBox.Text))
            {
                MessageBox.Show("Invalid path entered.", "Error");
                RefreshTextBox();
                return;
            }

            SetValue(textBox.Text);
        }
        #endregion

        #region Check and Get Value
        public string CheckAndGetPath(bool showMsg = false)
        {
            if (textBox.Text == "")
            {
                string msg = $"No folderpath for {attName} provided";
                if (showMsg)
                {
                    MessageBox.Show(msg, "Error");
                }
                throw new Exception(msg);
            }
            else if (!Directory.Exists(textBox.Text))
            {
                string msg = $"Invalid folder path for {attName}:\n{textBox.Text}";
                if (showMsg)
                {
                    MessageBox.Show(msg, "Error");
                }
                throw new Exception(msg);
            }
            return textBox.Text;
        }
        #endregion
    }

    public class FileTextBox : AttributeTextBox
    {
        #region Initialisation
        public Button setButton;
        public Button openButton;
        public string extension;
        public FileTextBox(string attName, TextBox textBox, Button button) : base(attName, textBox, false)
        {
            setButton = button;
            SubscribeToFileTextBoxEvents();
        }

        private void SubscribeToFileTextBoxEvents()
        {
            setButton.Click += new EventHandler(setButton_Click);
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        public void AddOpenButton(Button openButton, string extension = "")
        {
            this.openButton = openButton;
            if (extension != "")
            {
                this.extension = extension;
            }
            this.openButton.Click += new EventHandler(openButton_Click);
        }
        #endregion


        #region Basic Operations
        private void setButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileBrowserDialog = new OpenFileDialog();
            if (fileBrowserDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            SetValue(fileBrowserDialog.FileName);
        }

        protected new void textBox_LostFocus(object sender, EventArgs e)
        {
            SetValueFromTextBox();
        }

        protected new void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
        }

        protected new void SetValueFromTextBox(bool showMsg = false)
        {
            if (textBox.Text == "")
            {
                ResetValue();
                return;
            }

            if (!Path.IsPathRooted(textBox.Text) || !File.Exists(textBox.Text))
            {
                MessageBox.Show("Invalid path provided", "Error");
                RefreshTextBox();
                return;
            }

            SetValue(textBox.Text);
        }
        
        private void openButton_Click(object sender, EventArgs e)
        {
            if (textBox.Text == "") { return; }

            try
            {
                System.Diagnostics.Process.Start(textBox.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }
        #endregion

        #region Check and Get Value
        public string CheckAndGetValue(bool showMsg = false)
        {
            if (textBox.Text == "")
            {
                string msg = $"Invalid filepath for {attName}";
                if (showMsg)
                {
                    MessageBox.Show(msg, "Error");
                }
                throw new Exception(msg);
            }
            else if (!Path.IsPathRooted(textBox.Text) || !File.Exists(textBox.Text))
            {
                string msg = $"Invalid file path for {attName}:\n\n{textBox.Text}";
                if (showMsg)
                {
                    MessageBox.Show(msg, "Error");
                }
                throw new Exception(msg);
            }
            return textBox.Text;
        }
        #endregion

    }
    
    public class TargetCriteria
    {
        // Links together several objects to allow checking of source criterial easily
        private RangeTextBox criteriaSource;
        private AttributeTextBox criteriaValue;
        private ComboBoxAttribute logicSymbol;
        public bool isCriteriaDouble;

        public TargetCriteria(RangeTextBox criteriaSource, ComboBoxAttribute logicSymbol, AttributeTextBox criteriaValue)
        {
            this.criteriaSource = criteriaSource;
            this.criteriaValue = criteriaValue;
            this.logicSymbol = logicSymbol;
            CheckInputs();
        }

        public void CheckInputs()
        {
            // Criteria Source Range
            Range CriteriaSourceRange;
            try
            {
                CriteriaSourceRange = criteriaSource.GetRangeFromFullAddress();
            }
            catch
            {
                throw new Exception("Invalid input for Criteria Data Source.");
            }

            if (CriteriaSourceRange.Columns.Count != 1)
            {
                throw new Exception("Invalid input for Criteria Data Source, more than one column provided");
            }

            // LogicSymbol
            string logicSymbolString = logicSymbol.comboBox.Text;
            HashSet<string> validSymbols = new HashSet<string> { ">", "<", "=", ">=", "<=", "!=" };
            if (!validSymbols.Contains(logicSymbolString))
            {
                throw new Exception($"invalid logic symbol {logicSymbol}");
            }

            // CriteriaValue
            string CriteriaValue = criteriaValue.textBox.Text; ;
            if (CriteriaValue == "")
            {
                throw new Exception("Invalid input for Target Value (UR), field cannot be empty.");
            }

            // Try to parse UR into double
            isCriteriaDouble = false;
            if (double.TryParse(CriteriaValue, out double uRdouble))
            {
                isCriteriaDouble = true;
            }
        }

        public bool CriteriaMet()
        {
            string currentValue = criteriaSource.GetFirstCellContent();
            string targetValue = criteriaValue.textBox.Text;
            string logicSymbol = this.logicSymbol.comboBox.Text;
            
            if (isCriteriaDouble)
            {
                #region Convert numbers to double
                bool convertCurrent = double.TryParse(currentValue, out double currentValueDouble);
                bool convertTarget = double.TryParse(targetValue, out double targetValueDouble);
                #endregion

                switch (logicSymbol)
                {
                    case ">": return currentValueDouble > targetValueDouble;
                    case "<": return currentValueDouble < targetValueDouble;
                    case "=": return currentValueDouble == targetValueDouble;
                    case ">=": return currentValueDouble >= targetValueDouble;
                    case "<=": return currentValueDouble <= targetValueDouble;
                    case "!=": return currentValueDouble != targetValueDouble;
                    default: throw new Exception($"invalid logic symbol {logicSymbol}");
                }
            }
            else
            {
                switch (logicSymbol)
                {
                    case "=": return currentValue == targetValue;
                    case "!=": return currentValue != targetValue;
                    default: throw new Exception($"Invalid operator {logicSymbol} for comparing {currentValue} and {targetValue}");
                }
            }
        }
    }


}


