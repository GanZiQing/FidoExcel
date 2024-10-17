using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace ExcelAddIn2
{
    public partial class RangeSelector : Form
    {
        #region Initialise
        bool withSheet;
        MultipleRangeAttribute thisAtt;
        public RangeSelector(MultipleRangeAttribute thisAtt, bool withSheet)
        {
            this.thisAtt = thisAtt;
            this.withSheet = withSheet;
            InitializeComponent();
            FillListBox();
            CancelButton = cancelButton;
            SubscribeToEvents();

            #region Set Position
            StartPosition = FormStartPosition.CenterParent;
            #endregion

        }

        private void SubscribeToEvents()
        {
            rangeListBox.MouseDoubleClick += new MouseEventHandler(rangeListBox_DoubleClick);
            FormClosing += new FormClosingEventHandler((sender, e)=> setToNull());
        }

        private void FillListBox()
        {
            (string[] contents, _) = thisAtt.GetRanges();
            foreach(string item in contents)
            {
                rangeListBox.Items.Add(item);
            }
        }
        #endregion

        #region Add, Clear, Delete
        private void addRange_Click(object sender, EventArgs e)
        {
            object userInput = Globals.ThisAddIn.Application.InputBox("Select data range", "Select data", Type: 8);
            if (userInput is bool)
            {
                return;
            }
            string address = "";
            if (withSheet)
            {
                address += ((Range)userInput).Worksheet.Name + "!";
            }
            address += ((Range)userInput).Address[false, false];

            rangeListBox.Items.Add(address);
            this.Focus();
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!CommonUtilities.Confirmation("Clear all items in list?")) { return; }
                rangeListBox.Items.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!CommonUtilities.Confirmation("Delete selected items from list?")) { return; }
                for (int i = rangeListBox.SelectedIndices.Count - 1; i >= 0; i--)
                {
                    rangeListBox.Items.RemoveAt(rangeListBox.SelectedIndices[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }
        #endregion

        #region Editing Contents

        private void editButton_Click(object sender, EventArgs e)
        {
            foreach (int i in rangeListBox.SelectedIndices)
            {
                string currentValue = rangeListBox.Items[i].ToString();
                string currentAddress = "";
                try
                {
                    Range currentRange = CommonUtilities.CheckStringIsRange(currentValue, withSheet);
                    currentRange.Worksheet.Activate();
                    currentAddress = CommonUtilities.CheckStringIsRange(currentValue, withSheet).Address[false, false];
                }
                catch { }

                object userInput = Globals.ThisAddIn.Application.InputBox("Edit range", "Edit", currentAddress, Type: 8);
                if (userInput is bool)
                {
                    return;
                }

                else if (userInput is Range)
                {
                    string address = "";
                    if (withSheet)
                    {
                        address += ((Range)userInput).Worksheet.Name + "!";
                    }
                    address += ((Range)userInput).Address[false, false];
                    rangeListBox.Items[i] = address;
                }
            }
        }

        private void rangeListBox_DoubleClick(object sender, MouseEventArgs e)
        {
            int i = rangeListBox.IndexFromPoint(e.Location);

            // Ensure the index is valid
            if (i == System.Windows.Forms.ListBox.NoMatches)
            {
                return;
            }

            string currentValue = rangeListBox.Items[i].ToString();
            string newValue;
            using (CustomInputBox stringInputBox = new CustomInputBox(currentValue))
            {
                stringInputBox.Text = "Edit Range Value";
                stringInputBox.ShowDialog();

                if (stringInputBox.Value == null)
                {
                    return;
                }
                newValue = stringInputBox.Value;
            }

            try
            {
                CommonUtilities.CheckStringIsRange(newValue, withSheet);
                rangeListBox.Items[i] = newValue.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        #endregion

        #region Offset and Copy
        private void offSet_Click(object sender, EventArgs e)
        {
            offSetSelections(true);
        }

        int rowOffset = 0;
        int colOffset = 0;
        private void copy_Click(object sender, EventArgs e)
        {
            offSetSelections(false);
        }

        private void offSetSelections(bool toEdit)
        {
            try
            {
                if (useSameOffsetCheck.Checked)
                {
                    if (!GetOffsets())
                    {
                        return;
                    }
                }

                (int[] selectedIndices, string[] selectedItems) = GetSelectedIndices();
                List<int> highlightIndex = new List<int>();
                for (int i = selectedIndices.Length - 1; i >= 0; i--)
                {
                    int listBoxIndex = selectedIndices[i];
                    string currentValue = selectedItems[i];
                    Range currentRange = CommonUtilities.CheckStringIsRange(currentValue, withSheet);
                    if (!useSameOffsetCheck.Checked)
                    {
                        if (!GetOffsets(currentValue))
                        {
                            return;
                        }
                    }

                    Range newRange = currentRange.Offset[rowOffset, colOffset];
                    string address = "";
                    if (withSheet)
                    {
                        address += newRange.Worksheet.Name + "!";
                    }
                    address += newRange.Address[false, false];

                    if (toEdit)
                    {
                        rangeListBox.Items[listBoxIndex] = address;
                        highlightIndex.Add(listBoxIndex);
                    }
                    else
                    {
                        rangeListBox.Items.Add(address);
                        highlightIndex.Add(rangeListBox.Items.Count-1);
                    }
                }

                rangeListBox.ClearSelected();
                foreach (int i in highlightIndex)
                {
                    rangeListBox.SetSelected(i, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private bool GetOffsets(string address = null)
        {
            using (CustomInputBox stringInputBox = new CustomInputBox(rowOffset.ToString()))
            {
                stringInputBox.Text = $"Set x offset";
                if (address != null)
                {
                    stringInputBox.Instruction = $"Set X offset for {address}";
                }
                else
                {
                    stringInputBox.Instruction = $"Set X offset";
                }
                
                stringInputBox.type = "int";
                stringInputBox.ShowDialog();

                if (stringInputBox.Value == null)
                {
                    return false;
                }
                rowOffset = Int32.Parse(stringInputBox.Value);
            }

            using (CustomInputBox stringInputBox = new CustomInputBox(colOffset.ToString()))
            {
                stringInputBox.Text = $"Set Y offset";
                if (address != null)
                {
                    stringInputBox.Instruction = $"Set Y offset for {address}";
                }
                else
                {
                    stringInputBox.Instruction = $"Set Y offset";
                }
                stringInputBox.type = "int";
                stringInputBox.ShowDialog();

                if (stringInputBox.Value == null)
                {
                    return false;
                }
                colOffset = Int32.Parse(stringInputBox.Value);
            }
            return true;
        }
        #endregion

        #region Closing Buttons

        private void okButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Check Values
                List<string> checkedStrings = new List<string>();
                List<string> failedStrings = new List<string>();
                List<int> failedIndices = new List<int>();

                int i = 0;
                foreach (object obj in rangeListBox.Items)
                {
                    string item = obj.ToString();
                    try
                    {
                        CommonUtilities.CheckStringIsRange(item, withSheet);
                        checkedStrings.Add(item);
                    }
                    catch //(Exception ex)
                    {
                        failedStrings.Add(item);
                        failedIndices.Add(i);
                    }
                    i++;
                }

                // No remaining values
                if (failedStrings.Count == 0)
                {
                    thisAtt.SetByStringList(checkedStrings);
                    thisAtt.rangeSelector = null;
                    Close();
                    return;
                }

                // Delete missing values
                string msg = "The following ranges cannot be found, continue without these values?\n";
                foreach (string value in failedStrings)
                {
                    msg += $"{value}\n";
                }

                if (!CommonUtilities.Confirmation(msg, false))
                {
                    return;
                }

                thisAtt.SetByStringList(checkedStrings);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void setToNull()
        {
            thisAtt.rangeSelector = null;
        }
        #endregion

        #region Moving Buttons
        private void moveUp_Click(object sender, EventArgs e)
        {
            try
            {
                (int[] selectedIndices, string[] selectedItems) = GetSelectedIndices();
                HashSet<int> newSelectedIndices = new HashSet<int>();
                for (int i = selectedIndices.Length - 1; i >= 0; i--)
                {
                    if (selectedIndices[i] == 0 || newSelectedIndices.Contains(selectedIndices[i] - 1))
                    {
                        newSelectedIndices.Add(selectedIndices[i]);
                        continue;
                    }

                    rangeListBox.Items.RemoveAt(selectedIndices[i]);
                    rangeListBox.Items.Insert(selectedIndices[i] - 1, selectedItems[i]);
                    newSelectedIndices.Add(selectedIndices[i] - 1);
                }

                // Select all items
                rangeListBox.ClearSelected();
                foreach (int i in newSelectedIndices)
                {
                    rangeListBox.SetSelected(i, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void moveDown_Click(object sender, EventArgs e)
        {
            try
            {
                (int[] selectedIndices, string[] selectedItems) = GetSelectedIndices();
                HashSet<int> newSelectedIndices = new HashSet<int>();
                
                for (int i = 0; i < selectedIndices.Length; i++)
                {
                    if (selectedIndices[i] == rangeListBox.Items.Count - 1 || newSelectedIndices.Contains(selectedIndices[i] + 1))
                    {
                        newSelectedIndices.Add(selectedIndices[i]);
                        continue;
                    }

                    rangeListBox.Items.RemoveAt(selectedIndices[i]);
                    rangeListBox.Items.Insert(selectedIndices[i] + 1, selectedItems[i]);
                    newSelectedIndices.Add(selectedIndices[i] + 1);
                }

                // Select all items
                rangeListBox.ClearSelected();
                foreach (int i in newSelectedIndices)
                {
                    rangeListBox.SetSelected(i, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void moveToBottom_Click(object sender, EventArgs e)
        {
            try
            {
                (int[] selectedIndices, string[] selectedItems) = GetSelectedIndices();
                int insertIndex = rangeListBox.Items.Count - 1;
                HashSet<int> newSelectedIndices = new HashSet<int>();

                for (int i = 0; i < selectedIndices.Length; i++)
                {
                    rangeListBox.Items.RemoveAt(selectedIndices[i]);
                    rangeListBox.Items.Insert(insertIndex, selectedItems[i]);
                    insertIndex--;
                }

                // Select all items
                rangeListBox.ClearSelected();
                for (int i = insertIndex + 1; i < rangeListBox.Items.Count; i++)
                {
                    rangeListBox.SetSelected(i, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void moveToTop_Click(object sender, EventArgs e)
        {
            try
            {
                (int[] selectedIndices, string[] selectedItems) = GetSelectedIndices();
                HashSet<int> newSelectedIndices = new HashSet<int>();
                int newIndex = 0;
                for (int i = selectedIndices.Length - 1; i >= 0; i--)
                {
                    if (selectedIndices[i] == 0 || newSelectedIndices.Contains(selectedIndices[i] - 1))
                    {
                        newSelectedIndices.Add(selectedIndices[i]);
                        continue;
                    }

                    rangeListBox.Items.RemoveAt(selectedIndices[i]);
                    rangeListBox.Items.Insert(newIndex, selectedItems[i]);
                    newSelectedIndices.Add(newIndex);
                    newIndex++;
                }

                // Select all items
                rangeListBox.ClearSelected();
                foreach (int i in newSelectedIndices)
                {
                    rangeListBox.SetSelected(i, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private (int[] selectedIndices, string[] selectedItems) GetSelectedIndices()
        {
            List<string> selectedItems = new List<string>();
            List<int> selectedIndices = new List<int>();

            for (int i = rangeListBox.SelectedIndices.Count - 1; i >= 0; i--)
            {
                int index = rangeListBox.SelectedIndices[i];
                selectedIndices.Add(index);
                selectedItems.Add((string)rangeListBox.Items[index]);
            }
            return (selectedIndices.ToArray(), selectedItems.ToArray());
        }
        #endregion

    }
}
