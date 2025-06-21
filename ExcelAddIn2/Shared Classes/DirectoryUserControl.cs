using ExcelAddIn2.Excel_Pane_Folder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;
using Microsoft.VisualBasic.FileIO; // For Recycling

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
            #region Headers
            List<string> headers = new List<string> { "File Path", "Folder Name", "File Name" };
            AddHeaderMenuToButton(importFilePath, headers);
            AddHeaderMenuToButton(importSpecificFile, headers);

            headers = new List<string> { "File Path", "Parent Folder Name", "Folder Name" };
            AddHeaderMenuToButton(importFolderPath, headers);

            headers = new List<string> { "File Name" };
            AddHeaderMenuToButton(importFileName, headers);
            AddHeaderMenuToButton(importSpecificFileNames, headers);

            headers = new List<string> { "Folder Name" };
            AddHeaderMenuToButton(importFolderName, headers);

            headers = new List<string> { "File Path", "Folder", "File Name", "New File Name", "Status" };
            AddHeaderMenuToButton(renameFiles, headers);
            #endregion

            #region Get from Dialogue Box
            //Add File Details from Dialogue
            AddContextStripEvent(importFilePath, "Get From Dialogue Box", (sender, e) => importFilePath_Click(sender, e));
            AddContextStripEvent(importFileName, "Get From Dialogue Box", (sender, e) => importFileName_Click(sender, e));
            AddContextStripEvent(importFolderPath, "Get From Dialogue Box", (sender, e) => importFolderPath_Click(sender, e));
            AddContextStripEvent(importFolderName, "Get From Dialogue Box", (sender, e) => importFolderName_Click(sender, e));
            AddContextStripEvent(importSpecificFile, "Get from Dialogue Box", (sender, e) => importSpecificFile_Click(sender, e));
            AddContextStripEvent(importSpecificFileNames, "Get from Dialogue Box", (sender, e) => importSpecificFileNames_Click(sender, e));
            #endregion
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

            CheckBoxAttribute checkAtt = new CheckBoxAttribute("checkNested_Dir", nestedFoldersCheck, true);
            CustomAttributeDic.Add(checkAtt.attName, checkAtt);
            
            checkAtt = new CheckBoxAttribute("includeExt_Dir", addExtensionCheck, true);
            CustomAttributeDic.Add(checkAtt.attName, checkAtt);

            AttributeTextBox ExtensionType = new AttributeTextBox("ExtensionType", dispExtension, true);
            var thisCustomAtt = new CheckBoxAttribute("includeExtension", addExtensionCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            checkAtt = new CheckBoxAttribute("appendToFileName_Dir", appendFileNameCheck, false);
            CustomAttributeDic.Add(checkAtt.attName, checkAtt);

            checkAtt = new CheckBoxAttribute("mergeFolders_Dir", mergeFoldersCheck, false);
            CustomAttributeDic.Add(checkAtt.attName, checkAtt);
            #endregion

            attributeCreated = true;
        }

        private void AddToolTips()
        {
            #region Get
            toolTip1.SetToolTip(importFilePath,
                "For each file in selected folder, return:\n" +
                "  Full Directory | Folder Name | File Name");

            toolTip1.SetToolTip(importFileName,
                "For each file in selected folder, return:\n" +
                "  File Name");

            toolTip1.SetToolTip(importFolderPath,
                "For each folder in selected folder, return:\n" +
                "  Full Directory | Parent Folder Name | Folder Name");

            toolTip1.SetToolTip(importFolderName,
                "For each folder in selected folder, return:\n" +
                "  Folder Name");

            toolTip1.SetToolTip(dispExtension,
                "Provide extension of the file type to limit search to. Extension should start with \".\"\n" +
                "  Case sensitive (aka .cbd =/= .CBD)\n" +
                "  Use comma to define multiple file types (e.g \".pdf, .xlsx\")\n" +
                "  Use \".excel\" to filter for .xlsx, .xlsm, .xlsb, .xls, .csv");

            toolTip1.SetToolTip(importSpecificFile,
                "For each file in selected folder that matches the specified extension, return:\n" +
                "  Full Directory | Folder Name | File Name");

            toolTip1.SetToolTip(importSpecificFileNames,
                "For each file in selected folder that matches the specified extension, return:\n" +
                "  File Name");
            #endregion

            #region Edit
            toolTip1.SetToolTip(renameFiles,
                "Rename files. Takes input:\n" +
                "  Path | *Parent Folder | *File/Folder Name | New File/Folder Name \n" +
                "  *Data in these columns are not used.");

            toolTip1.SetToolTip(copyFiles,
                "Copy files from the selected file paths to target folder. Takes input:\n" +
                "  File Path\n" +
                "Primary usage is to use excel to filter desired files before copying.");

            toolTip1.SetToolTip(moveFiles,
                "Move files from the selected file paths to target folder. Takes input:\n" +
                "  File Path\n" +
                "Primary usage is to use excel to filter desired files before copying.");

            toolTip1.SetToolTip(createFolders,
                "Create folders. Takes input:\n" +
                "  Relative Folder Path\n");
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
            createFolders.Enabled = false;
            copyFiles.Enabled = false;
            moveFiles.Enabled = false;

            this.Height = 224;
        }
        public void ShowUpToSpecifyExtension()
        {
            renameFiles.Enabled = false;
            createFolders.Enabled = false;
            copyFiles.Enabled = false;
            moveFiles.Enabled = false;

            this.Height = 428;
        }
        #endregion
        #endregion

        #region Directory Management

        #region Buttons
        private void importFilePath_Click(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem)
            {
                string destinationFolder;
                try
                {
                    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                    customFolderBrowser.ShowDialog();
                    destinationFolder = customFolderBrowser.GetFolderPath();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                importFilesOrFolders(true, false, overwriteDirectory: destinationFolder);
            }
            if (sender is System.Windows.Forms.Button)
            {
                importFilesOrFolders(true, false);
            }

        }

        private void importFileName_Click(object sender, EventArgs e)
        {

            if (sender is ToolStripMenuItem)
            {
                string destinationFolder;
                try
                {
                    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                    customFolderBrowser.ShowDialog();
                    destinationFolder = customFolderBrowser.GetFolderPath();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                importFilesOrFolders(true, true, overwriteDirectory: destinationFolder);
            }
            if (sender is System.Windows.Forms.Button)
            {
                importFilesOrFolders(true, true);
            }

            //importFilesOrFolders(true, true);
        }

        private void importFolderPath_Click(object sender, EventArgs e)
        {
            //importFilesOrFolders(false, false);

            if (sender is ToolStripMenuItem)
            {
                string destinationFolder;
                try
                {
                    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                    customFolderBrowser.ShowDialog();
                    destinationFolder = customFolderBrowser.GetFolderPath();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                importFilesOrFolders(false, false, overwriteDirectory: destinationFolder);
            }
            if (sender is System.Windows.Forms.Button)
            {
                importFilesOrFolders(false, false);
            }
        }

        private void importFolderName_Click(object sender, EventArgs e)
        {
            //importFilesOrFolders(false, true);
            if (sender is ToolStripMenuItem)
            {
                string destinationFolder;
                try
                {
                    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                    customFolderBrowser.ShowDialog();
                    destinationFolder = customFolderBrowser.GetFolderPath();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                importFilesOrFolders(false, true, overwriteDirectory: destinationFolder);
            }
            if (sender is System.Windows.Forms.Button)
            {
                importFilesOrFolders(false, true);
            }
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

            //importFilesOrFolders(true, false, dispExtension.Text);
            if (sender is ToolStripMenuItem)
            {
                string destinationFolder;
                try
                {
                    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                    customFolderBrowser.ShowDialog();
                    destinationFolder = customFolderBrowser.GetFolderPath();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                importFilesOrFolders(true, false, dispExtension.Text, overwriteDirectory: destinationFolder);
            }
            if (sender is System.Windows.Forms.Button)
            {
                importFilesOrFolders(true, false, dispExtension.Text);
            }
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

            //importFilesOrFolders(true, true, dispExtension.Text);
            if (sender is ToolStripMenuItem)
            {
                string destinationFolder;
                try
                {
                    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                    customFolderBrowser.ShowDialog();
                    destinationFolder = customFolderBrowser.GetFolderPath();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                importFilesOrFolders(true, true, dispExtension.Text, overwriteDirectory: destinationFolder);
            }
            if (sender is System.Windows.Forms.Button)
            {
                importFilesOrFolders(true, true, dispExtension.Text);
            }
        }
        #endregion

        #region Main File Path Function
        private void importFilesOrFolders(bool isFile, bool nameOnly, string specifiedExtension = "", string overwriteDirectory = "")
        {
            try
            {
                #region Set Diriectory
                string directoryPath;
                if (overwriteDirectory == "")
                {
                    CheckDirectory();
                    directoryPath = dispDirectory.Text;
                }
                else
                {
                    directoryPath = overwriteDirectory;
                }
                #endregion

                #region Get Parameters
                Workbook activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Worksheet activeSheet = activeBook.ActiveSheet;
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                #endregion

                #region Get Files or Folders
                List<string> files = new List<string>();
                if (isFile)
                {
                    getFiles(directoryPath, ref files, nestedFoldersCheck.Checked, specifiedExtension);
                }
                else { getFolders(directoryPath, ref files, nestedFoldersCheck.Checked); }

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

                if (!nameOnly) { WriteToExcelSelectionAsRow(0, 0, true, fullPath, folderName, fileName); }
                else { WriteToExcelSelectionAsRow(0, 0, true, fileName); }
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

        #region Create Folder
        private void createFolders_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not Implemented");
        }
        #endregion

        #region Copy
        private void copyFiles_Click(object sender, EventArgs e) { copyOrMoveFiles(true); }
        private void moveFiles_Click(object sender, EventArgs e) { copyOrMoveFiles(false); }

        private void copyOrMoveFiles(bool isCopy)
        {
            try
            {
                #region Check Input Size
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                try { CheckRangeSize(selectedRange, 0, 1); }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }

                string[] filePaths = GetContentsAsStringArray(selectedRange, true);
                if (filePaths == null) { throw new Exception("No file paths provided in selected range"); }
                List<string> failedPaths = new List<string>();

                foreach (string filePath in filePaths)
                {
                    if (!File.Exists(filePath) & !Directory.Exists(filePath)) { failedPaths.Add(filePath); }
                }

                if (failedPaths.Count > 0)
                {
                    string msg = "The following file path(s) do not exist:\n";
                    foreach (string failedPath in failedPaths)
                    {
                        msg += failedPath + "\n";
                    }
                    throw new Exception(msg);
                }
                #endregion

                #region Get Destination
                CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                customFolderBrowser.ShowDialog();
                string destinationFolder = customFolderBrowser.GetFolderPath();
                #endregion

                #region Copy Files
                int filesCopied = 0;
                foreach (string filePath in filePaths)
                {
                    #region File
                    if (File.Exists(filePath)) // Is a File
                    {
                        string destinationPath = Path.Combine(destinationFolder, Path.GetFileName(filePath));
                        if (File.Exists(destinationPath))
                        {
                            if (appendFileNameCheck.Checked) // Try new names
                            {
                                string originalFileName = Path.GetFileNameWithoutExtension(filePath);
                                string originalExtension = Path.GetExtension(filePath);

                                int attempt = 1;
                                int maxAttempt = 100;
                                while (attempt < maxAttempt)
                                {
                                    string newFileName = originalFileName + $" ({attempt})" + originalExtension;
                                    string newDestinationPath = Path.Combine(destinationFolder, newFileName);
                                    if (File.Exists(newDestinationPath)) { attempt++; continue; } // Skip remaining and try new name

                                    if (isCopy) { File.Copy(filePath, newDestinationPath, false); }
                                    else { File.Move(filePath, newDestinationPath); }
                                    filesCopied++;
                                    break;
                                }

                                if (attempt >= maxAttempt)
                                {
                                    MessageBox.Show($"Unable to find new name for {Path.GetFileName(filePath)}. File skipped.\n" +
                                        $"File Path: {filePath}");
                                }
                            }
                            else // Ask to delete
                            {
                                DialogResult res = MessageBox.Show($"File {Path.GetFileName(filePath)} already exist at destination, overwrite?", "Warning", MessageBoxButtons.YesNoCancel);
                                if (res == DialogResult.Cancel) { throw new Exception("Terminated by user"); }
                                else if (res == DialogResult.Yes)
                                {
                                    if (isCopy) { File.Copy(filePath, destinationPath, true); }
                                    else
                                    {
                                        File.Delete(destinationPath);
                                        File.Move(filePath, destinationPath);
                                    }
                                    filesCopied++;
                                }
                            }                           
                        }
                        else
                        {
                            if (isCopy) { File.Copy(filePath, destinationPath, false); }
                            else { File.Move(filePath, destinationPath); }
                            filesCopied++;
                        }
                    }
                    #endregion

                    #region Directory
                    else // Is a Directory (case that file and folder does not exist is already taken care of at the start
                    {
                        string destinationPath = Path.Combine(destinationFolder, Path.GetFileName(filePath));
                        if (Directory.Exists(destinationPath))
                        {
                            if (mergeFoldersCheck.Checked) { throw new NotImplementedException("Merge Folders not implemented"); }
                            else // Ask to delete
                            {
                                string folderName = new DirectoryInfo(destinationPath).Name;
                                DialogResult res = MessageBox.Show($"Folder {folderName} already exist at destination, delete existing?", "Warning", MessageBoxButtons.YesNoCancel);
                                if (res == DialogResult.Cancel) { throw new Exception("Terminated by user"); }
                                else if (res == DialogResult.Yes)
                                {
                                    FileSystem.DeleteDirectory(destinationPath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);

                                    if (isCopy) { copyFolder(filePath, destinationPath); }
                                    else { Directory.Move(filePath, destinationPath); }
                                    filesCopied++;
                                }
                            }
                        }
                        else
                        {
                            if (isCopy) { copyFolder(filePath, destinationPath); }
                            else { Directory.Move(filePath, destinationPath); }
                            filesCopied++;
                        }
                    }
                    #endregion
                }
                #endregion

                MessageBox.Show($"{filesCopied}/{filePaths.Length} files copied to new directory.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void copyFolder(string sourcePath, string destinationPath)
        {
            // This function assumes both path have been checked and are ok
            // Written by chatGPT

            // Create destination directory if it doesn't exist
            if (!Directory.Exists(destinationPath))
            {
                Directory.CreateDirectory(destinationPath);
            }

            // Copy all files
            foreach (string filePath in Directory.GetFiles(sourcePath))
            {
                string fileName = Path.GetFileName(filePath);
                string destFile = Path.Combine(destinationPath, fileName);
                File.Copy(filePath, destFile, true); // Overwrite if exists
            }

            // Recursively copy all subdirectories
            foreach (string dirPath in Directory.GetDirectories(sourcePath))
            {
                string dirName = Path.GetFileName(dirPath);
                string destSubDir = Path.Combine(destinationPath, dirName);
                copyFolder(dirPath, destSubDir);
            }
        }
        #endregion


    }
}
