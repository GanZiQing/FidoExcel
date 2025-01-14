using ExcelAddIn2.Excel_Pane_Folder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;

namespace ExcelAddIn2
{
    public partial class DirectoryUserControl : UserControl
    {
        #region Init
        public DirectoryUserControl()
        {
            InitializeComponent();
            AddToolTips();
            AddContextStrips();
        }

        private void AddContextStrips()
        {
            List<string> headers = new List<string> { "File Path", "Folder Name", "File Name"};
            AddHeaderMenuToButton(importFilePath, headers);

            headers = new List<string> { "File Path", "Folder Name"};
            AddHeaderMenuToButton(importFolderPath, headers);

            headers = new List<string> { "File Name" };
            AddHeaderMenuToButton(importFileName, headers);

            headers = new List<string> { "Folder Name" };
            AddHeaderMenuToButton(importFolderName, headers);

            headers = new List<string> { "File Path", "Folder", "File Name", "New File Name", "Status" };
            AddHeaderMenuToButton(renameFiles, headers);
        }

        bool attributeCreated = false;
        public void CreateAttributes(ref
            Dictionary<string, AttributeTextBox> AttributeTextBoxDic, ref
            Dictionary<string, CustomAttribute> CustomAttributeDic)
        {
            if (attributeCreated) { return; }
            #region Directory
            DirectoryTextBox FolderPath = new DirectoryTextBox("FolderPath", dispDirectory, setDirectory);
            FolderPath.AddOpenButton(dirOpenPath);
            AttributeTextBoxDic.Add("FolderPath", FolderPath);
            AttributeTextBox ExtensionType = new AttributeTextBox("ExtensionType", dispExtension, true);
            var thisCustomAtt = new CheckBoxAttribute("includeExtension", addExtensionCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            #endregion
            attributeCreated = true;
        }

        private void AddToolTips()
        {
            #region Directory
            toolTip1.SetToolTip(importFilePath,
                "For each file in selected folder, return:\n" +
                "Full Directory | Folder Name | Filename");

            toolTip1.SetToolTip(renameFiles,
                "Rename files assuming the selected range (4 columns) of the following format:" +
                "File Path | Folder | File Name | File Name\n" +
                "Data in Folder and Origional File name columns are not used.");

            toolTip1.SetToolTip(mergeFolders,
                "Inserts reference headers used for \"Import Paths\" and \"Rename Files\"\n" +
                "File Path | Folder | File Name | New File Name | Status");
            #endregion
        }


        #region Resize
        public void ShowFileDetailsOnly()
        {
            
            importFolderPath.Enabled = false;
            importFolderName.Enabled = false; 
            dispExtension.Enabled = false;
            importSpecificFile.Enabled = false;
            importSpecificFileNames.Enabled = false;
            renameFiles.Enabled = false;
            mergeFolders.Enabled = false;
            this.Height = 224;
        }
        public void ShowUpToSpecifyExtension()
        {
            renameFiles.Enabled = false;
            mergeFolders.Enabled = false;
            this.Height = 428;
        }
        #endregion
        #endregion

        #region Directory Management

        #region Buttons
        private void importFilePath_Click(object sender, EventArgs e)
        {
            importFilesOrFolders(true, false);
        }

        private void importFileName_Click(object sender, EventArgs e)
        {
            importFilesOrFolders(true, true);
        }
        
        private void importFolderPath_Click(object sender, EventArgs e)
        {
            importFilesOrFolders(false, false);
            //CheckDirectory();

            //#region Get Parameters
            //string directoryPath = dispDirectory.Text;
            //Workbook activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //Worksheet activeSheet = activeBook.ActiveSheet;
            //Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            //#endregion

            //#region Call method to get files and folders
            //List<string> directories = new List<string>();
            //getFolderDirectories(directoryPath, ref directories, checkNestedFolders.Checked);
            //#endregion

            //#region Print results
            //// Print files array
            //string[] folder_name = new string[directories.Count()];
            //string[] full_path = new string[directories.Count()];
            //int i = 0;
            //foreach (string folder in directories)
            //{
            //    full_path[i] = folder;
            //    //folder_name[i] = Path.GetFileName(Path.GetDirectoryName(file));
            //    folder_name[i] = Path.GetFileName(folder);
            //    i++;
            //}
            //try
            //{
            //    WriteToExcel(0, 0, true, full_path, folder_name);
            //}
            //catch (Exception ex)
            //{
            //    if (ex.Message == "Nothing found to print")
            //    {
            //        MessageBox.Show("No results found");
            //    }
            //    else
            //    {
            //        MessageBox.Show($"Error encountered\n\n{ex.Message}");
            //    }
            //}
            //#endregion
        }
        
        private void importFolderName_Click(object sender, EventArgs e)
        {
            importFilesOrFolders(false, true);
        }
        
        private void importSpecificFile_Click(object sender, EventArgs e)
        {
            #region Get Extension
            if (dispExtension.Text == "")
            {
                MessageBox.Show("No extension type provided.", "Error");
                return;
            }
            else if (dispExtension.Text[0] != '.')
            {
                MessageBox.Show($"Invalid extension type provided. Extension should start with '.'", "Error");
                return;
            }
            #endregion

            importFilesOrFolders(true, false, dispExtension.Text);
        }
        private void importSpecificFileNames_Click(object sender, EventArgs e)
        {
            #region Get Extension
            if (dispExtension.Text == "")
            {
                MessageBox.Show("No extension type provided.", "Error");
                return;
            }
            else if (dispExtension.Text[0] != '.')
            {
                MessageBox.Show($"Invalid extension type provided. Extension should start with '.'", "Error");
                return;
            }
            #endregion

            importFilesOrFolders(true, true, dispExtension.Text);
        }
        #endregion
        
        #region Main File Path Function
        private void importFilesOrFolders(bool isFile, bool nameOnly, string specifiedExtension = "")
        {
            try
            {
                #region Get Parameters
                CheckDirectory();
                string directoryPath = dispDirectory.Text;
                Workbook activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Worksheet activeSheet = activeBook.ActiveSheet;
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                #endregion

                #region Get Directories and Files
                List<string> files = new List<string>();
                if (isFile)
                {
                    if (specifiedExtension == "")
                    {
                        getFiles(directoryPath, ref files, checkNestedFolders.Checked);
                    }
                    else
                    {
                        getFiles(directoryPath, ref files, checkNestedFolders.Checked, specifiedExtension);
                    }

                }
                else { getFolders(directoryPath, ref files, checkNestedFolders.Checked); }

                if (files.Count == 0) { throw new Exception("No items found to print"); }
                #endregion

                #region Print results
                string[] fullPath = files.ToArray();
                string[] folderName = new string[files.Count()];
                string[] fileName = new string[files.Count()];

                for (int i = 0; i < fullPath.Length; i++)
                {
                    string file = fullPath[i];
                    if (addExtensionCheck.Checked) { fileName[i] = Path.GetFileName(file); }
                    else { fileName[i] = Path.GetFileNameWithoutExtension(file); }
                    folderName[i] = Path.GetFileName(Path.GetDirectoryName(file));
                }

                if (isFile && !nameOnly) { WriteToExcelSelectionAsRow(0, 0, true, fullPath, folderName, fileName); }
                else if (isFile && nameOnly) { WriteToExcelSelectionAsRow(0, 0, true, fileName); }
                else if (!isFile && !nameOnly) { WriteToExcelSelectionAsRow(0, 0, true, fullPath, folderName); }
                else if (!isFile && nameOnly) { WriteToExcelSelectionAsRow(0, 0, true, folderName); }
                #endregion

                #region Format Path to be less annoying
                if (!nameOnly)
                {
                    Range startCell = selectedRange.Cells[1];
                    Range endCell = startCell.Offset[fullPath.Length - 1];
                    Range formatRange = selectedRange.Worksheet.Range[startCell, endCell];
                    formatRange.Cells.Font.Color = Color.Gainsboro;
                }
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #endregion

        #region Helper Functions
        private void CheckDirectory()
        {
            if (dispDirectory.Text == "")
            {
                throw new ArgumentException("No folder path provided");
            }
            else if (!Directory.Exists(dispDirectory.Text))
            {
                throw new ArgumentException($"Invalid folder path:\n{dispDirectory.Text}");
            }
        }
        private static void WriteToExcel(int rowOff, int colOff, bool setCellToText, params Array[] arrays)
        {
            // This code takes any number of arrays (of various types) and outputs them into excel 
            // Output order depends on order of the input array
            // Output location is the first cell of the current selection, offset by rowOff and colOff

            // Find number of rows and columns
            int numRow = 0;
            int numCol = arrays.Length;
            for (int col = 0; col < arrays.Length; col++)
            {
                if (arrays[col].Length > numRow)
                {
                    numRow = arrays[col].Length; // Finds max number of rows out of all the various arrays
                }
            }

            #region Check if data exist
            if (numRow == 0)
            {
                throw new Exception("Nothing found to print");
            }
            #endregion

            #region Set Excel Params
            // Add section to read input data from Excel
            Workbook activeWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet activeWorkSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            // Write to Excel
            activeWB.Application.ScreenUpdating = false;

            // Write the entire array to the worksheet in one go using Value2
            Range startCell = activeWorkSheet.Cells[selectedRange.Row + rowOff, selectedRange.Column + colOff];
            Range endCell = startCell.Offset[numRow - 1, numCol - 1];
            Range writeRange = activeWorkSheet.Range[startCell, endCell];
            #endregion

            #region Set cell formatting to text
            if (setCellToText)
            {
                for (int col = 0; col < arrays.Length; col++)
                {
                    if (arrays[col] is string[])
                    {
                        Range locStartCell = startCell.Offset[0, col];
                        Range locEndCell = locStartCell.Offset[numRow - 1, 0];
                        Range formatCell = activeWorkSheet.Range[locStartCell, locEndCell];
                        formatCell.NumberFormat = "@";
                    }
                }
            }
            #endregion

            // Initiate object
            object[,] dataArray = new object[numRow, numCol];
            for (int col = 0; col < arrays.Length; col++)
            {
                for (int row = 0; row < arrays[col].Length; row++)
                {
                    dataArray[row, col] = arrays[col].GetValue(row);
                }
            }

            writeRange.Value2 = dataArray;

            activeWB.Application.ScreenUpdating = true;
            activeWorkSheet = null;
        }

        #endregion
        #endregion

        #region Rename
        private void renameFiles_Click(object sender, EventArgs e)
        {
            #region Check Input Size
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            try { CheckRangeSize(selectedRange, 0, 4); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }
            #endregion

            #region Get Confirmation
            if (DialogResult.OK != MessageBox.Show($"Confirm to rename {selectedRange.Rows.Count} files? This cannot be undone.", "Confirmation"))
            {
                return;
            }
            #endregion

            #region Read Excel Info
            ExcelTable thisTable = new ExcelTable(selectedRange, "Selected Table");
            thisTable.AddColumn(1, "sourcePaths");
            thisTable.AddColumn(4, "newNames");
            thisTable.ReadRangeFromTable();

            string[] sourcePaths = thisTable.GetColumnFromName("sourcePaths").ConvertRangeToStringArray();
            string[] newNames = thisTable.GetColumnFromName("newNames").ConvertRangeToStringArray();
            #endregion

            #region Change Names
            string[] status = new string[sourcePaths.Length];
            int failures = 0;
            for (int i = 0; i < sourcePaths.Length; i++)
            {
                try
                {
                    string sourcePath = sourcePaths[i];
                    string newName = newNames[i];
                    if (newName == "")
                    {
                        throw new Exception("Error: File Name cannot be empty");
                    }
                    string folder = Path.GetDirectoryName(sourcePath);
                    string newPath = Path.Combine(folder, newName);

                    status[i] = renameOneFile(sourcePath, newPath);
                }
                catch (Exception ex)
                {
                    status[i] = "Error: " + ex.Message;
                }
                if (status[i] != "Completed: File renamed")
                {
                    failures++;
                }
            }
            #endregion

            if (failures == 0)
            {
                MessageBox.Show("Rename operation completed.\n" +
                     $"{sourcePaths.Length - failures}/{sourcePaths.Length} files renamed", "Completed");
            }
            else
            {
                CommonUtilities.WriteToExcelSelectionAsRow(0, 4, false, status);
                MessageBox.Show("Rename operation incomplete.\n" +
                     $"{sourcePaths.Length - failures}/{sourcePaths.Length} files renamed. Check status.", "Completed");
            }
        }
        
        private string renameOneFile(string sourcePath, string newPath)
        {

            FileAttributes attribute = File.GetAttributes(sourcePath);

            if (attribute == FileAttributes.Directory)
            {
                try
                {
                    Directory.Move(sourcePath, newPath);
                    return "Completed: File renamed";
                }
                catch (Exception ex)
                {
                    return "Error: " + ex.Message;
                }
            }
            else
            {
                #region Check if Path Exist
                if (!File.Exists(sourcePath))
                {
                    //MessageBox.Show($"The following file does not exist\n{sourcePath}", "Error");
                    //throw new Exception($"The following file does not exist\n{sourcePath}", "Error");
                    return "Error: File does not exist";
                }
                #endregion

                #region Check Extension
                if (!Path.HasExtension(newPath))
                {
                    newPath += Path.GetExtension(sourcePath);
                }
                else if (Path.GetExtension(sourcePath) != Path.GetExtension(newPath))
                {
                    //MessageBox.Show("Inconsistent extension type.\n" +
                    //    $"Original extension is {Path.GetExtension(sourcePath)} but new extension is {Path.GetExtension(newPath)}.\n" +
                    //    "Source Path:\n" +
                    //    $"{sourcePath}");
                    return "Warning: Inconsistent extension type";
                }
                #endregion

                try
                {
                    File.Move(sourcePath, newPath);
                    return "Completed: File renamed";
                }
                catch (Exception ex)
                {
                    return "Error: " + ex.Message;
                }
            }
        }

        #endregion
    }
}
