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
using ListBox = System.Windows.Forms.ListBox;

namespace ExcelAddIn2
{
    public partial class SheetSelector : Form
    {
        #region Initializer
        //public string attName;
        MultipleSheetsAttribute thisAttribute;
        public SheetSelector(MultipleSheetsAttribute thisAttribute)
        {
            //this.attName = attName;
            this.thisAttribute = thisAttribute;
            InitializeComponent();
            InitializeListBox();
            StartPosition = FormStartPosition.CenterParent;
            CancelButton = MyCancelButton;
        }

        private void InitializeListBox()
        {
            // Get all sheets and sheets saved in attribute
            Sheets allSheets = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;
            // Get all sheets that was saved previously
            HashSet<string> extractedSheetNames = thisAttribute.GetSheetNamesHash();
            List<string> availbleSavedSheetNames = new List<string>();

            if (extractedSheetNames.Count == 0)
            {
                foreach (Worksheet sheet in allSheets)
                {
                    LeftListBox.Items.Add(sheet.Name);
                }
            }
            else
            {
                //Sort sheets into respective boxes
                foreach (Worksheet sheet in allSheets)
                {
                    if (extractedSheetNames.Contains(sheet.Name))
                    {
                        // Add to right sheetbox
                        RightListBox.Items.Add(sheet.Name);
                        extractedSheetNames.Remove(sheet.Name);
                        availbleSavedSheetNames.Add(sheet.Name);
                    }
                    else
                    {
                        // Add to left sheet box
                        LeftListBox.Items.Add(sheet.Name);
                    }
                }

                //If there's any remaining sheets in the saved sheet names, reset names
                if (extractedSheetNames.Count > 0)
                {
                    thisAttribute.SetSheetsByName(availbleSavedSheetNames);
                    string msg = "The following sheet(s) are not found and have been removed from selection:\n";
                    foreach (string sheet in extractedSheetNames)
                    {
                        msg += sheet + "\n";
                    }
                    MessageBox.Show(msg, "Sheet(s) not found.");
                }
            }
        }

        private void RefreshListBox()
        {

        }
        #endregion

        #region Move Sheet Buttons
        #region Helper Functions
        private void MoveAllItems(ListBox source, ListBox destination)
        {
            // Check if source and destination are not null
            if (source == null || destination == null)
                throw new ArgumentNullException("ListBoxes cannot be null.");

            // Add all items from source to destination
            destination.Items.AddRange(source.Items);
            // Clear all items from the source listBox
            source.Items.Clear();
        }

        private void MoveSelectedItems(ListBox source, ListBox destination)
        {
            // Check if source and destination are not null
            if (source == null || destination == null)
                throw new ArgumentNullException("ListBoxes cannot be null.");

            // Collect the selected items in an array
            var selectedItems = new object[source.SelectedItems.Count];
            source.SelectedItems.CopyTo(selectedItems, 0);

            // Add selected items to the destination
            destination.Items.AddRange(selectedItems);

            // Remove selected items from the source
            foreach (var item in selectedItems)
            {
                source.Items.Remove(item);
            }
        }
        #endregion

        private void MoveAllRight_Click(object sender, EventArgs e)
        {
            MoveAllItems(LeftListBox, RightListBox);
        }

        private void MoveSelectionRight_Click(object sender, EventArgs e)
        {
            MoveSelectedItems(LeftListBox, RightListBox);
        }

        private void MoveSelectionLeft_Click(object sender, EventArgs e)
        {
            MoveSelectedItems(RightListBox, LeftListBox);
        }

        private void MoveAllLeft_Click(object sender, EventArgs e)
        {
            MoveAllItems(RightListBox, LeftListBox); 
        }
        #endregion

        #region Confirmation Buttons

        #region Helper Functions
        private List<string> GetListBoxItems(ListBox listBox)
        {
            List<string> items = new List<string>();
            foreach (var item in listBox.Items)
            {
                items.Add(item.ToString());
            }
            return items;
        }

        private List<string> GetListBoxSelectedItems(ListBox listBox)
        {
            List<string> items = new List<string>();
            foreach (var item in listBox.SelectedItems)
            {
                items.Add(item.ToString());
            }
            return items;
        }
        #endregion
        private void ConfirmationButton_Click(object sender, EventArgs e)
        {
            // Set Sheets
            List<string> sheetsToSave = GetListBoxItems(RightListBox);
            thisAttribute.SetSheetsByName(sheetsToSave);
            Close();
        }

        private void MyCancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            List<string> selectedSheets = GetListBoxItems(RightListBox);
            // Check Input and Warning
            if (selectedSheets.Count == 0)
            {
                MessageBox.Show("Please select sheets to delete.");
                return;
            }

            string msg = "Delete the following sheets? This cannot be undone.\n";
            foreach (string sheet in selectedSheets)
            {
                msg += sheet + "\n";
            }
            DialogResult response = MessageBox.Show(msg, "Confirmation", MessageBoxButtons.YesNo);
            if (response == DialogResult.No) { return; }
            // Delete Files
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                foreach (string sheet in selectedSheets)
                {
                    Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheet].Delete();
                    RightListBox.Items.Remove(sheet);
                }
            }
            catch
            {
                MessageBox.Show("Error deleting sheets.","Error");
            }
            finally
            {
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
        }
        #endregion

        #region Other Functions
        public void AddToDictionary(Dictionary<string,CustomAttribute> attributeDict)
        {
            attributeDict[thisAttribute.attName] = thisAttribute;
        }

        public void ShowDeleteSheet()
        {
            DeleteButton.Visible = true;
            ConfirmationButton.Visible = false;
            DeleteButton.Location = ConfirmationButton.Location;
        }
        #endregion
    }

    //class DeleteSheetSelector : SheetSelector
    //{
    //    public DeleteSheetSelector(string attName) : base(attName)
    //    {
            
    //    }

    //    private void SetDeleteButton()
    //    {
    //        ConfirmationButton.Text = "Delete";
    //    }
    //}
}
