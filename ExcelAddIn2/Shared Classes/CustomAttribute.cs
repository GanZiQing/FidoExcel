using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;

namespace ExcelAddIn2
{
    public class CustomAttribute
    {
        public string attName;
        public string defaultValue;
        public object attValue;

        public CustomAttribute(string AttName, string defaultValue = null)
        {
            this.attName = AttName;
            this.defaultValue = defaultValue;
            // Get value if it exist
            (bool success , string value) = GetStringValueFromProp();
            if (success)
            {
                attValue = value;
            }
            else
            {
                attValue = defaultValue;
            }
        }

        public (bool, DocumentProperty) GetProp()
        {
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

        public (bool, string) GetStringValueFromProp()
        {
            DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
            foreach (DocumentProperty prop in AllCustProps)
            {
                if (prop.Name == attName)
                {
                    return (true, prop.Value.ToString());
                }
            }
            return (false, "");
        }

        public bool SetValue(object inputAttValue, bool showMsg = false)
        {
            try
            {
                DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
                (bool exist, DocumentProperty prop) = GetProp();
                if (exist)
                {
                    if (prop.Value.ToString() == inputAttValue.ToString())
                    {
                        return true;
                    }
                    prop.Value = inputAttValue;
                    attValue = inputAttValue;
                    if (showMsg)
                    {
                        MessageBox.Show(attName + " updated to " + inputAttValue.ToString());
                    }
                }
                else
                {
                    AllCustProps.Add(attName, false, MsoDocProperties.msoPropertyTypeString, inputAttValue.ToString());
                    attValue = inputAttValue;
                    if (showMsg)
                    {
                        MessageBox.Show(attName + " added as " + inputAttValue.ToString());
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding value {attName}\n\n" + ex.Message);
                return false;
            }
        }

        public virtual void ResetValue()
        {
            if (defaultValue == null)
            {
                (bool gotValue, DocumentProperty thisProp) = GetProp();
                if (gotValue)
                {
                    thisProp.Delete();
                    attValue = null;
                }
            }
            else
            {
                SetValue(defaultValue);
            }
        }

        public virtual bool ImportValue(string Value)
        {
            bool success = SetValue(Value, false);
            return success;
        }
    }

    public class ComboBoxAttribute : CustomAttribute
    {
        public ComboBox comboBox;
        bool isInitialisation = true;
        public ComboBoxAttribute(string AttName, ComboBox comboBox,string defaultValue = null) : base(AttName, defaultValue)
        {
            this.comboBox = comboBox;
            comboBox.MouseWheel += new MouseEventHandler(PreventScroll);
            comboBox.SelectedIndexChanged += new EventHandler(comboBox_SelectedIndexChanged);
            RefreshComboBox();
        }

        #region Event Handlers
        private void PreventScroll(object sender, MouseEventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }

        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isInitialisation)
            {
                isInitialisation = false; // Skip setting value if we are initialising
            }
            else
            {
                try
                {
                    SetComboValueFromBox();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error setting combobox value");
                }
            }
            
        }
        #endregion

        public void SetComboValueFromBox()
        {
            SetValue(comboBox.Text, false);
        }

        public void RefreshComboBox()
        {
            (bool hasValue, DocumentProperty thisProp) = GetProp();
            if (hasValue)
            {
                comboBox.Text = thisProp.Value;
            }
            else
            {
                if (defaultValue != null)
                {
                    comboBox.Text = defaultValue;
                    //SetComboValue();
                }
            }
        }

        #region Import export
        public override void ResetValue()
        {
            if (defaultValue == null)
            {
                (bool gotValue, DocumentProperty thisProp) = GetProp();
                thisProp.Delete();
            }
            else
            {
                SetValue(defaultValue);
            }
            RefreshComboBox();
        }

        public override bool ImportValue(string Value)
        {
            bool success = SetValue(Value);
            RefreshComboBox();
            return success;
        }
        #endregion
    }

    public class MultipleSheetsAttribute : CustomAttribute
    {
        #region Initialise
        Button launchButton;
        bool isDel;
        public MultipleSheetsAttribute(string attName, Button launchButton, bool isDel = false, string defaultValue = null) : base(attName, defaultValue)
        {
            this.launchButton = launchButton;
            SubscribeToButton();
            this.isDel = isDel;
        }

        private void SubscribeToButton()
        {
            launchButton.Click += new EventHandler(launchButton_Click);
        }

        private void launchButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (SheetSelector sheetSelector = new SheetSelector(this))
                {
                    if (isDel)
                    {
                        sheetSelector.ShowDeleteSheet();
                    }
                    sheetSelector.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }
        #endregion

        public HashSet<Worksheet> GetSheets()
        {
            string[] sheetNames = GetSheetNamesArray();

            HashSet<Worksheet> allSheets = new HashSet<Worksheet>();
            foreach (string sheetName in sheetNames)
            {
                allSheets.Add(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName]);
            }
            return allSheets;
        }

        public HashSet<string> GetSheetNamesHash()
        {
            string[] sheetNames = GetSheetNamesArray();
            HashSet<string> allSheetName = new HashSet<string>();
            foreach (string sheetName in sheetNames)
            {
                allSheetName.Add(sheetName);
            }
            return allSheetName;
        }

        public string[] GetSheetNamesArray()
        {
            if (attValue == null)
            {
                return new string[0];
            }
            string[] sheetNames = (attValue.ToString()).Split('*');

            List<string> existingSheetList = new List<string>();
            List<string> missingSheetList = new List<string>();
            foreach (string sheetName in sheetNames)
            {
                try
                {
                    Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    existingSheetList.Add(sheetName);
                }
                catch
                {
                    missingSheetList.Add(sheetName);
                }
            }

            if (missingSheetList.Count > 0)
            {
                string msg = "The following sheets cannot be found. Remove sheets from selection and continue?\n";
                foreach (string sheet in missingSheetList)
                {
                    msg += sheet + "\n";
                }

                DialogResult result = MessageBox.Show(msg, "Confirmation", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                {
                    throw new Exception("Terminated by user");
                }

                SetSheetsByName(existingSheetList);
                return existingSheetList.ToArray();
            }
            else
            {
                return sheetNames;
            }            
        }

        public void SetSheetsByWorksheet(List<Worksheet> worksheets)
        {
            if (worksheets.Count == 0)
            {
                ResetValue();
                return;
            }

            string attValue = "";
            foreach (Worksheet sheet in worksheets)
            {
                attValue+= "*" + sheet.Name;
            }
            attValue = attValue.Substring(1);

            SetValue(attValue);
            return;
        }

        public void SetSheetsByName(List<string> worksheets)
        {
            if (worksheets.Count == 0)
            {
                ResetValue();
                return;
            }

            string attValue = "";
            if (worksheets.Count > 0)
            {
                foreach (string sheet in worksheets)
                {
                    attValue += "*" + sheet;
                }
                attValue = attValue.Substring(1);
            }
            SetValue(attValue);
        }
    }

    public class MultipleRangeAttribute : CustomAttribute
    {
        #region Initialise
        Button launchButton;
        bool withSheet;
        public MultipleRangeAttribute(string attName, Button launchButton, bool withSheet) : base(attName)
        {
            this.launchButton = launchButton;
            this.withSheet = withSheet;
            SubscribeToButton();
        }

        private void SubscribeToButton()
        {
            launchButton.Click += new EventHandler(launchButton_Click);
        }

        public RangeSelector rangeSelector = null;
        private void launchButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (rangeSelector == null)
                {
                    rangeSelector = new RangeSelector(this, withSheet);
                    IntPtr excelHandle = new IntPtr(Globals.ThisAddIn.Application.Hwnd);
                    NativeWindow excelWindow = new NativeWindow();
                    excelWindow.AssignHandle(excelHandle);
                    rangeSelector.Show(excelWindow);
                }
                else
                {
                    rangeSelector.Activate();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }
        #endregion

        public (string[] rangeStrings, Range[] ranges) GetRanges()
        {
            if (attValue == null)
            {
                return (new string[0], new Range[0]);
            }

            string[] rangeAddresses = (attValue.ToString()).Split('*');

            List<string> existingAddresses = new List<string>();
            List<string> missingAddresses = new List<string>();
            List<Range> ranges = new List<Range>();
            foreach (string rangeString in rangeAddresses)
            {
                try
                {
                    Range thisRange = CommonUtilities.CheckStringIsRange(rangeString, withSheet);
                    ranges.Add(thisRange);
                    existingAddresses.Add(rangeString);
                }
                catch
                {
                    missingAddresses.Add(rangeString);
                }
            }

            if (missingAddresses.Count == 0)
            {
                return (rangeAddresses, ranges.ToArray());
            }

            string msg = "The following ranges cannot be found. Remove range from selection and continue?\n";
            foreach (string range in missingAddresses)
            {
                msg += range + "\n";
            }
                
            bool toContinue = CommonUtilities.Confirmation(msg, false);
            if (toContinue)
            {
                SetByStringList(existingAddresses);
                return (existingAddresses.ToArray(), ranges.ToArray());
            }
            else
            {
                return (rangeAddresses.ToArray(), ranges.ToArray());
            }
            
        }

        public void SetByStringList(List<string> items)
        {
            // Probably can generalise this for MutlipleSheetsAttribute too
            if (items.Count == 0)
            {
                ResetValue();
                return;
            }

            string attValue = "";
            if (items.Count > 0)
            {
                foreach (string sheet in items)
                {
                    attValue += "*" + sheet;
                }
                attValue = attValue.Substring(1);
            }
            SetValue(attValue);
        }
    }

    public class FontDialogAttribute
    {
        Button setButton;
        FontDialog fontDialog;
        public FontDialogAttribute(string baseAttName, Button button)
        {
            CustomAttribute colorAtt = new CustomAttribute(baseAttName + "_color");
            CustomAttribute fontSizeAtt = new CustomAttribute(baseAttName + "_fontSize");
            CustomAttribute fontTypeAtt = new CustomAttribute(baseAttName + "_fontType");
            setButton = button;
            fontDialog = new FontDialog
            {
                ShowColor = true
            };
            RefreshFontDialog();
        }

        private void SubscribeToFontEvents()
        {
            setButton.Click += new EventHandler(setButton_Click);
        }

        private void setButton_Click(object sender, EventArgs e)
        {     
            if (fontDialog.ShowDialog() != DialogResult.OK)
            {
                SetAttributes();
            }
        }

        private void RefreshFontDialog()
        {
            //(bool hasColour, string color) GetStringValue();
            
            //fontDialog.Font = new Font("Arial", 16);
        }

        private void SetAttributes()
        {

        }
    }

    public class CheckBoxAttribute: CustomAttribute
    {
        #region Init
        bool defaultState;
        CheckBox checkBox;
        public CheckBoxAttribute(string attName, CheckBox checkBox, bool defaultState = false) : base (attName)
        {
            this.defaultState = defaultState;
            this.checkBox = checkBox;
            SubscribeToEvents();
            RefreshCheckBox();
        }
        private void SubscribeToEvents()
        {
            checkBox.CheckedChanged += new EventHandler(SetBooleanValueFromCheckBox);
        }

        private void SetBooleanValueFromCheckBox(object sender, EventArgs e)
        {
            SetBooleanValue(checkBox.Checked);
        }
        #endregion

        #region Set Value and Refresh
        private bool SetBooleanValue(bool state)
        {
            (bool exist, DocumentProperty prop) = GetProp();
            if (exist && prop.Type == MsoDocProperties.msoPropertyTypeBoolean)
            {
                if (prop.Value == state)
                {
                    return true;
                }
                prop.Value = state;
            }
            else if (exist && prop.Type != MsoDocProperties.msoPropertyTypeBoolean)
            {
                prop.Delete();
                DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
                AllCustProps.Add(attName, false, MsoDocProperties.msoPropertyTypeBoolean, state);
            }
            else
            {
                DocumentProperties AllCustProps = Globals.ThisAddIn.Application.ActiveWorkbook.CustomDocumentProperties;
                AllCustProps.Add(attName, false, MsoDocProperties.msoPropertyTypeBoolean, state);
            }
            return true;
        }

        public void RefreshCheckBox()
        {
            try
            {
                (bool hasValue, DocumentProperty thisProp) = GetProp();
                if (hasValue)
                {
                    checkBox.Checked = thisProp.Value;
                }
                else
                {
                    checkBox.Checked = defaultState;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in refreshing check box");
                checkBox.Checked = defaultState;
            }
        }
        #endregion

        #region Import Export
        public override void ResetValue()
        {
            (bool gotValue, DocumentProperty thisProp) = GetProp();
            if (gotValue)
            {
                thisProp.Delete();
            }
            RefreshCheckBox();
        }

        public override bool ImportValue(string Value)
        {
            Boolean.TryParse(Value, out bool state);
            bool success = SetBooleanValue(state);
            RefreshCheckBox();
            return success;
        }
        #endregion
    }
}

