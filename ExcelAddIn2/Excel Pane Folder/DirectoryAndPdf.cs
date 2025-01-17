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
using System.Threading;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using static ExcelAddIn2.CommonUtilities;
using TextBox = System.Windows.Forms.TextBox;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.CompilerServices;
using System.Linq.Expressions;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using System.Runtime.InteropServices;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class DirectoryAndPdf : UserControl
    {
        #region Initialisers
        Workbook ThisWorkBook;
        Application ThisApplication;
        DocumentProperties AllCustProps;
        Dictionary<string, AttributeTextBox> AttributeTextBoxDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> CustomAttributeDic = new Dictionary<string, CustomAttribute>();

        public DirectoryAndPdf()
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
            List<string> headers = new List<string>
            {
            "Section Title",
            "Start Pg Num",
            "End Pg Num",
            "Total Pg Num",
            "Insert New Page",
            "File Path",
            };
            AddHeaderMenuToButton(advanceMerge, headers);

            headers = new List<string> { "Output File Name - leave blank for default", "Sheet Name - leave blank to print all", "File Path" };
            AddHeaderMenuToButton(printWorkbooks, headers);
        }

        private void CreateAttributes()
        {
            // Create Attribute Objects 
            #region Print
            // Print
            AttributeTextBox PrintFolder1 = new AttributeTextBox("PrintFolder", DispPrintFolder, true);
            PrintFolder1.type = "partial filepath";
            AttributeTextBoxDic.Add("PrintFolder", PrintFolder1);
            AttributeTextBox AppLeft1 = new AttributeTextBox("AppLeft", DispAppLeft, true);
            AppLeft1.type = "filename";
            AttributeTextBoxDic.Add("AppLeft", AppLeft1);
            AttributeTextBox AppRight1 = new AttributeTextBox("AppRight", DispAppRight, true);
            AppRight1.type = "filename";
            AttributeTextBoxDic.Add("AppRight", AppRight1);
            MultipleSheetsAttribute SavedPrintSheet = new MultipleSheetsAttribute("SavedPrintSheet", SetSheetsToPrint);
            CustomAttributeDic.Add("SavedPrintSheet", SavedPrintSheet);
            #endregion

            #region Print Workbooks
            AttributeTextBox thisAtt = new DirectoryTextBox("destFolder_printPDF", dispDestFolder, setDestFolder);
            ((DirectoryTextBox)thisAtt).AddOpenButton(openDestFolder);
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);
            CustomAttribute thisCustomAtt = new CheckBoxAttribute("overwriteDest_printPDF", overwritePrintPath, false);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            #endregion

            #region Directory
            //DirectoryTextBox FolderPath = new DirectoryTextBox("FolderPath", dispDirectory, setDirectory);
            //FolderPath.AddOpenButton(dirOpenPath);
            //AttributeTextBoxDic.Add("FolderPath", FolderPath);
            //AttributeTextBox ExtensionType = new AttributeTextBox("ExtensionType", dispExtension, true);
            //thisCustomAtt = new CheckBoxAttribute("includeExtension", addExtensionCheck, true);
            //CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            directoryUserControl.CreateAttributes(ref AttributeTextBoxDic, ref CustomAttributeDic);
            #endregion

            #region Compare Folders
            DirectoryTextBox dir = new DirectoryTextBox("FolderPath1", dispFolder1, setFolder1);
            dir.AddOpenButton(openFolder1);
            AttributeTextBoxDic.Add(dir.attName, dir);

            dir = new DirectoryTextBox("FolderPath2", dispFolder2, setFolder2);
            dir.AddOpenButton(openFolder2);
            AttributeTextBoxDic.Add(dir.attName, dir);

            thisAtt = new AttributeTextBox("ExtensionTypeComparison", dispExtension, true);
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisCustomAtt = new CheckBoxAttribute("IncludeExtensionComparison", specifyExtensionCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisCustomAtt = new CheckBoxAttribute("SearchSubFoldersComparison", searchSubFoldersCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            #endregion

            #region PDF
            DirectoryTextBox PdfFolderPath = new DirectoryTextBox("PdfFolderPath", dispPdfOutFolder, setPdfOutFolder);
            PdfFolderPath.AddOpenButton(openPdfOutFolder);
            AttributeTextBoxDic.Add("PdfFolderPath", PdfFolderPath);

            AttributeTextBox MergeName = new AttributeTextBox("MergeName", dispMergeName, true);
            MergeName.type = "filename";
            AttributeTextBoxDic.Add("MergeName", MergeName);

            thisCustomAtt = new CheckBoxAttribute("createBookmarks", createBookmarksCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            FileTextBox RefTitlePageFile = new FileTextBox("RefTitlePageFile", dispRefTitlePage, setRefTitlePage);
            AttributeTextBoxDic.Add("RefTitlePageFile", RefTitlePageFile);

            AttributeTextBox TitleFontSize = new AttributeTextBox("TitleFontSize", dispTitleFontSize, true);
            TitleFontSize.type = "int";
            TitleFontSize.SetDefaultValue("25");
            AttributeTextBoxDic.Add("TitleFontSize", TitleFontSize);

            thisAtt = new AttributeTextBox("PdfOpenDelay", dispOpenDelay, true);
            thisAtt.type = "double";
            thisAtt.SetDefaultValue("0.1");
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region Add Page Number

            AttributeTextBox FirstPageNum = new AttributeTextBox("FirstPageNum ", dispFirstPageNum, true);
            FirstPageNum.type = "int";
            FirstPageNum.SetDefaultValue("1");
            AttributeTextBoxDic.Add("FirstPageNum", FirstPageNum);

            AttributeTextBox IgnorePageNum = new AttributeTextBox("IgnorePageNum ", dispSkipPage, true);
            IgnorePageNum.type = "int";
            IgnorePageNum.SetDefaultValue("0");
            AttributeTextBoxDic.Add("IgnorePageNum", IgnorePageNum);

            AttributeTextBox AppendFileName = new AttributeTextBox("AppendFileName ", dispAppendName, true);
            AppendFileName.type = "filename";
            AppendFileName.SetDefaultValue("_addPg");
            AttributeTextBoxDic.Add("AppendFileName", AppendFileName);

            AttributeTextBox PageNumFontSize = new AttributeTextBox("PageNumFontSize ", dispFontSize, true);
            PageNumFontSize.type = "int";
            PageNumFontSize.SetDefaultValue("8");
            AttributeTextBoxDic.Add("PageNumFontSize", PageNumFontSize);

            AttributeTextBox OffsetX = new AttributeTextBox("OffsetX", dispOffsetX, true);
            OffsetX.type = "double";
            OffsetX.SetDefaultValue("20");
            AttributeTextBoxDic.Add("OffsetX", OffsetX);

            AttributeTextBox OffsetY = new AttributeTextBox("OffsetY", dispOffsetY, true);
            OffsetY.type = "double";
            OffsetY.SetDefaultValue("20");
            AttributeTextBoxDic.Add("OffsetY", OffsetY);
            #endregion

            #region Drafting Sheet Number
            thisAtt = new AttributeTextBox("thisSheetX_sheetRenum", dispThisSheetX);
            thisAtt.SetDefaultValue("30");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("thisSheetY_sheetRenum", dispThisSheetY);
            thisAtt.SetDefaultValue("725");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("totalSheetX_sheetRenum", dispTotalSheetX);
            thisAtt.SetDefaultValue("30");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("totalSheetY_sheetRenum", dispTotalSheetY);
            thisAtt.SetDefaultValue("750");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("totalSheetNum_sheetRenum", dispTotalSheetY); // Need to sort this out
            thisAtt.type = "int";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("fontSize_sheetRenum", dispFontSizeSheetNum);
            thisAtt.SetDefaultValue("12"); 
            thisAtt.type = "int";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisCustomAtt = new CheckBoxAttribute("renameFile_sheetRenum", renameFilesCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisCustomAtt = new CheckBoxAttribute("addSheetNum_sheetRenum", addSheetNumberCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            #endregion
        }

        private void AddToolTips()
        {
            #region Print Tools
            toolTip1.SetToolTip(PrintCurrentSheet,
                "If folder name is empty, print files will be saved in current excel file path.\n" +
                "If folder name is provided, files are saved in a folder at the current excel file path.");
            #endregion

            #region Directory
            //toolTip1.SetToolTip(importFilePath,
            //    "For each file in selected folder, return:\n" +
            //    "Full Directory | Folder Name | Filename");

            //toolTip1.SetToolTip(renameFiles,
            //    "Rename files assuming the selected range (4 columns) of the following format:"+
            //    "File Path | Folder | File Name | File Name\n" +
            //    "Data in Folder and Origional File name columns are not used.");

            //toolTip1.SetToolTip(insertRenameHeader,
            //    "Inserts reference headers used for \"Import Paths\" and \"Rename Files\"\n" +
            //    "File Path | Folder | File Name | New File Name | Status");
            #endregion

            #region Compare Folders
            toolTip1.SetToolTip(unionFiles, "Union");
            toolTip1.SetToolTip(removeIntersectFiles, "Remove Intersect");
            toolTip1.SetToolTip(intersectFiles, "Intersect");
            toolTip1.SetToolTip(subtractFiles, "Subtract");
            toolTip1.SetToolTip(reverseSubtractFiles, "Reverse Subtract");
            #endregion

            #region Merge PDF
            toolTip1.SetToolTip(setRefTitlePage,
                "Set reference title page (provide PDF file).\n" +
                "Required for Advance Merge and Create Section Title Page");

            toolTip1.SetToolTip(basicMergePDF,
                "Merges PDF based on file path provided by current excel selection (1 col).\n" +
                "Output folder and file name as defined in task pane above.");
            toolTip1.SetToolTip(setRefTitlePage,
                "Sets the reference file for section headers.\n" +
                "This page will be inserted when \"Insert New Page\" is set to yes.\n" +
                "If left blank, an empty page will be used instead.");
            toolTip1.SetToolTip(generateSections,
                "Generates one page per section title provided.\n" +
                "Uses current Excel selection (2 cols).\n" +
                "Col 1 = Pdf Name, Col 2 = Section Title\n" +
                "Files saved in output folder defined above.");
            toolTip1.SetToolTip(advanceMerge,
                "Merge files and adds section dividers\nReference range based on current Excel selection (6 cols).\n" +
                "Assumes the following column order:\n" +
                "Section Title | Start Pg No. | End Pg No.| Total Pg No. | Insert New Page | File Path");
            toolTip1.SetToolTip(addPageNum,
                "Prompts user to select files to add page numeber to.\n" +
                "Ouput file will be saved in the same directory as original file with appended file name.\n" +
                "Note that this operation cannot be undone.");
            #endregion
        }
        #endregion

        #region Print
        #region Helper Functions
        private bool IsFileWritable(string filePath)
        {
            try
            {
                // Check if the file is open by another process
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // If no exception is thrown, the file is not in use by another process
                    return true;
                }
            }
            catch (IOException)
            {
                // The file is in use by another process
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                // The file is write-protected
                return false;
            }
        }

        private (bool, bool) CheckPrintPath(string folderPath, string fileName)
        {
            // Returns false if error is encountered
            //
            string pdfFilePath = Path.Combine(folderPath, fileName);
            bool toProceed = true;
            bool toTerminateAll = false;
            if (!Directory.Exists(folderPath))
            {
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch
                {
                    MessageBox.Show($"Unable to create folder at {folderPath}");
                    toProceed = false;
                    return (toProceed, toTerminateAll);
                }
            }

            // Check if file already exist
            if (File.Exists(pdfFilePath))
            {
                DialogResult result = MessageBox.Show($"Existing file found at {pdfFilePath}, overwrite file?", "File already exist", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.No)
                {
                    toProceed = false;
                    return (toProceed, toTerminateAll);
                }
                else if (result == DialogResult.Cancel)
                {
                    toProceed = false;
                    toTerminateAll = true;
                    return (toProceed, toTerminateAll);
                }
                else
                {
                    // Check if file is open
                    if (!IsFileWritable(pdfFilePath))
                    {
                        MessageBox.Show("Unable to overwrite file, check if file is open.", "Failed to overwrite");
                        toProceed = false;
                        return (toProceed, toTerminateAll);
                    }
                }
            }

            return (toProceed, toTerminateAll);
        }
        
        private void CreateDestinationFolder(string folderPath)
        {
            //Check if path exist
            if (!Directory.Exists(folderPath))
            {
                DialogResult result = MessageBox.Show("Folder does not currently exist. Create new folder?", "Error Opening Folder", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch
                    {
                        MessageBox.Show($"Unable to create folder at {folderPath}");
                        return;
                    }
                }
                else
                {
                    throw new Exception("Terminated by user");
                }
            }
        }
        #endregion

        #region Open Print Folder
        private void OpenPrintFolder_Click(object sender, EventArgs e)
        {
            try
            {
                // Get Path
                string workbookPath = Path.GetDirectoryName(ThisApplication.ActiveWorkbook.FullName);
                string folderPath = Path.Combine(workbookPath, DispPrintFolder.Text);
                //Check if path exist
                CreateDestinationFolder(folderPath);
                System.Diagnostics.Process.Start(folderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
        }
        #endregion

        #region Print Single
        private void PrintCurrentSheet_Click(object sender, EventArgs e)
        {
            // Get the path of the workbook
            Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sheetToPrint = ThisApplication.ActiveSheet;
            string workbookPath = Path.GetDirectoryName(activeWorkbook.FullName);

            try
            {
                // Get the name of the active worksheet
                string sheetName = sheetToPrint.Name;
                string fileName = DispAppLeft.Text + sheetName + DispAppRight.Text + ".pdf";
                // Create the full path for the PDF file
                string folderPath = Path.Combine(workbookPath, DispPrintFolder.Text);
                string pdfFilePath = Path.Combine(folderPath, fileName);

                (bool toProceed, bool toTerminateAll) = CheckPrintPath(folderPath, fileName);
                if (!toProceed)
                {
                    return;
                }
                if (toTerminateAll)
                {
                    return;
                }

                // Print
                if (PrintRangeCheck.Checked)
                {
                    ThisApplication.ActiveWindow.Selection.ExportAsFixedFormat(
                        XlFixedFormatType.xlTypePDF,
                        pdfFilePath,
                        XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: false,
                        FixedFormatExtClassPtr: Type.Missing);
                }
                else
                {
                    sheetToPrint.ExportAsFixedFormat(
                        XlFixedFormatType.xlTypePDF,
                        pdfFilePath,
                        XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: false,
                        FixedFormatExtClassPtr: Type.Missing);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to print\n\n" + ex.ToString(), "Failed to print");
                return;
            }
            MessageBox.Show($"Sheet printed.", "Completed");
        }
        #endregion

        #region Print Selected Sheets
        private void PrintSelSheets_Click(object sender, EventArgs e)
        {
            // Get the path of the workbook
            Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            string workbookPath = Path.GetDirectoryName(activeWorkbook.FullName);

            // Get names of sheets to print
            MultipleSheetsAttribute SavedPrintSheet = (MultipleSheetsAttribute) CustomAttributeDic["SavedPrintSheet"];
            //MutlipleSheetsAttribute SavedPrintSheet = new MutlipleSheetsAttribute("SavedPrintSheet");
            HashSet<string> PrintSheets = SavedPrintSheet.GetSheetNamesHash();
            int numPrintedSheets = 0;
            int totalNumSheets = PrintSheets.Count;


            foreach (string sheet in PrintSheets)
            {
                try
                {
                    // Get the name of the active worksheet
                    Worksheet sheetToPrint;
                    try
                    {
                        sheetToPrint = ThisApplication.ActiveWorkbook.Sheets[sheet];
                    }
                    catch
                    {
                        DialogResult result = MessageBox.Show($"Sheet {sheet} not found, continue to next sheet? \n Choose Cancel to cancel all remaining prints (if any).", "Error printing", MessageBoxButtons.YesNoCancel);
                        if (result == DialogResult.No)
                        {
                            continue;
                            //return;
                        }
                        else if (result == DialogResult.Cancel)
                        {
                            MessageBox.Show($"{numPrintedSheets} / {totalNumSheets} sheet(s) printed.", "Completed");
                            return;
                        }
                    }
                    sheetToPrint = ThisApplication.ActiveWorkbook.Sheets[sheet];

                    string sheetName = sheetToPrint.Name;
                    string fileName = DispAppLeft.Text + sheetName + DispAppRight.Text + ".pdf";
                    // Create the full path for the PDF file
                    string folderPath = Path.Combine(workbookPath, DispPrintFolder.Text);
                    string pdfFilePath = Path.Combine(folderPath, fileName);

                    (bool toProceed, bool toTerminateAll) = CheckPrintPath(folderPath, fileName);
                    if (!toProceed)
                    {
                        continue;
                    }
                    if (toTerminateAll)
                    {
                        MessageBox.Show($"{numPrintedSheets} / {totalNumSheets} sheet(s) printed.", "Completed");
                        return;
                    }

                    // Print
                    sheetToPrint.ExportAsFixedFormat(
                        XlFixedFormatType.xlTypePDF,
                        pdfFilePath,
                        XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: false,
                        FixedFormatExtClassPtr: Type.Missing);
                    numPrintedSheets += 1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to print" + ex.Message, "Failed to print");
                }
            }

            MessageBox.Show($"{numPrintedSheets} / {totalNumSheets} sheet(s) printed.", "Completed");
        }
        #endregion

        #region Advance Print Sheets
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
                CommonUtilities.WriteToExcelSelectionAsRow(0, 0, true, names.ToArray());
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
            #region Read Selection

            #endregion

            #region Rename Sheets

            #endregion
        }

        private void PrintSelSheetsAdvance_Click(object sender, EventArgs e)
        {
            #region Read Excel
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            Range nameRange = selectedRange.Rows[1];
            Range sheetsRange = selectedRange.Resize[selectedRange.Rows.Count - 1].Offset[1, 0];
            string baseName = ThisWorkBook.Name;
            string workbookPath = Path.GetDirectoryName(ThisWorkBook.FullName);
            string folderPath = Path.Combine(workbookPath, DispPrintFolder.Text);

            try
            {
                CreateDestinationFolder(folderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
            #endregion

            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                #region Print 
                try
                {
                    ThisApplication.ScreenUpdating = false;
                    for (int colNum = 1; colNum <= sheetsRange.Columns.Count; colNum++)
                    {
                        #region Print indivduals
                        
                        Range col = sheetsRange.Columns[colNum];
                        Sheets sheetsToPrint = ThisApplication.Sheets;
                        List<string> printedPaths = new List<string>();
                        foreach (Range cell in col.Cells)
                        {
                            if (cell.Value2 == null)
                            {
                                continue;
                            }
                            string sheetName = cell.Value2.ToString();
                            progressTracker.UpdateStatus($"Printing {sheetName}");
                            Worksheet sheetToPrint = ThisWorkBook.Sheets[sheetName];
                            string fileName = ThisWorkBook.Name + "_" + sheetName;
                            PrintSingleSheet(sheetToPrint, fileName, folderPath);
                            printedPaths.Add(Path.Combine(folderPath, fileName + ".pdf"));
                        }
                        #endregion

                        #region Merge and save final file
                        progressTracker.UpdateStatus($"Merging files for column {colNum}/{sheetsRange.Columns.Count}");
                        if (printedPaths.Count == 0)
                        {
                            continue;
                        }

                        string outputFileName;
                        if (nameRange.Cells[1, colNum].Value2 == null)
                        {
                            outputFileName = $"Column {colNum}.pdf";
                        }
                        else
                        {
                            outputFileName = nameRange.Cells[1, colNum].Value2.ToString() + ".pdf";
                        }

                        string outputPath = Path.Combine(folderPath, outputFileName);
                        MergeFiles(printedPaths, outputPath, true);
                        
                        worker.ReportProgress(ConvertToProgress(colNum, sheetsRange.Columns.Count));
                        #endregion

                        if (worker.CancellationPending) 
                        {
                            break; 
                        }
                    }
                }
                finally
                {
                    ThisApplication.ScreenUpdating = true;
                }
                #endregion
            });
        }
        
        private void GetAndPrintSingleSheet(Workbook workbook, string sheetName, string fileName, string folderPath)
        {
            Worksheet sheetToPrint;
            #region Print Sheet
            try
            {
                sheetToPrint = workbook.Sheets[sheetName];
            }
            catch //(Exception ex)
            {
                throw new ArgumentException($"Sheet {sheetName} not found in {workbook.Name}");
            }
            #endregion
            PrintSingleSheet(sheetToPrint, fileName, folderPath);
        }

        private void PrintSingleSheet(Worksheet sheet, string fileName, string folderPath)
        {
            try
            {
                (bool toProceed, bool toTerminateAll) = CheckPrintPath(folderPath, fileName);
                string pdfFilePath = Path.Combine(folderPath, fileName);

                if (!toProceed || toTerminateAll)
                {
                    throw new Exception("Terminated by user");
                }

                // Print
                sheet.ExportAsFixedFormat(
                        XlFixedFormatType.xlTypePDF,
                        pdfFilePath,
                        XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: false,
                        FixedFormatExtClassPtr: Type.Missing);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to print\n\n" + ex.Message);
            }
        }

        private void MergeFiles(List<string> filePaths, string outputPath, bool deleteOriginal = false)
        {
            PdfDocument outputDocument = new PdfDocument();
            foreach (string filepath in filePaths)
            {
                PdfDocument inputDocument = PdfReader.Open(filepath, PdfDocumentOpenMode.Import);
                for (int index = 0; index < inputDocument.PageCount; index++)
                {
                    PdfPage page = inputDocument.Pages[index];
                    outputDocument.AddPage(page);
                }
            }
            if (outputDocument.PageCount == 0)
            {
                MessageBox.Show("Final PDF file is empty, no file generated", "Error");
                return;
            }
            outputDocument.Save(outputPath);

            if (deleteOriginal)
            {
                foreach (string filepath in filePaths)
                {
                    try
                    {
                        File.Delete(filepath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Unable to delete file {filepath}" + ex.Message, "Error");
                    }
                }
            }
        }
        #endregion

        #region Print Workbooks
        private void printWorkbooks_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                #region Check Input Size
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                try { CheckRangeSize(selectedRange, 0, 3); }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }
                //string[] fileNames = GetContentsAsStringArray(selectedRange.Columns[1].Cells, false);
                string[] fileNames = GetAndCheckExcelFileNames(selectedRange.Columns[1].Cells, ".pdf");
                string[] sheetNames = GetContentsAsStringArray(selectedRange.Columns[2].Cells, false);
                string[] filePaths = GetContentsAsStringArray(selectedRange.Columns[3].Cells, false);
                #endregion

                #region Check validity of inputs
                foreach (string filePath in filePaths)
                {
                    if (filePath == "")
                    {
                        continue;
                    }

                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show($"FilePath provided below does not exist.\n{filePath}", "Error");
                        return;
                    }
                }
                #endregion

                #region Print
                try
                {
                    ThisApplication.ScreenUpdating = false;
                    ThisApplication.DisplayAlerts = false;
                    int numPrinted = 0;

                    #region Get Folder Path If Required
                    string folderPath = "";
                    if (!overwritePrintPath.Checked) 
                    { 
                        folderPath = ((DirectoryTextBox)AttributeTextBoxDic["destFolder_printPDF"]).CheckAndGetPath(); 
                    }
                    #endregion

                    Workbook workbookToPrint = null;
                    for (int rowNum = 0; rowNum < filePaths.Length; rowNum++)
                    {
                        string filePath = filePaths[rowNum];
                        if (filePath == "") { continue; }

                        #region Get Folder Path
                        if (overwritePrintPath.Checked)
                        {
                            folderPath = Path.GetDirectoryName(filePath);
                        }

                        #endregion

                        progressTracker.UpdateStatus($"Printing {filePath}");

                        string sheetName = sheetNames[rowNum];    
                        try
                        {
                            #region Get Workbook
                            if (workbookToPrint == null)
                            {
                                workbookToPrint = ThisApplication.Workbooks.Open(filePath);
                            }
                            else if (workbookToPrint.Path != filePath)
                            {
                                workbookToPrint.Close();
                                workbookToPrint = ThisApplication.Workbooks.Open(filePath);
                            }
                            string workbookFileName = Path.GetFileNameWithoutExtension(filePath);

                            #endregion

                            #region Print Sheets
                            string pdfFileName = GetPdfName(); ;
                            if (sheetName == "") // Print all visible workbook
                            {
                                //workbookFileName += ".pdf";
                                //pdfFileName = GetPdfName();
                                //PrintEntireWorkbook(workbookToPrint, sheetName, workbookFileName, folderPath);
                                PrintEntireWorkbook(workbookToPrint, sheetName, pdfFileName, folderPath);
                            }
                            else // Print single sheet
                            {
                                //workbookFileName += $"_{sheetName}.pdf";
                                //GetAndPrintSingleSheet(workbookToPrint, sheetName, workbookFileName, folderPath);
                                GetAndPrintSingleSheet(workbookToPrint, sheetName, pdfFileName, folderPath);
                            }
                            #endregion

                            numPrinted++;
                            worker.ReportProgress(ConvertToProgress(rowNum+1, filePaths.Length));

                            #region GetPDFName
                            string GetPdfName()
                            {
                                string returnPdfName;
                                if (fileNames[rowNum] != "")
                                {
                                    returnPdfName = fileNames[rowNum];
                                    if (!Path.HasExtension(returnPdfName)) // Extension type is checked earlier
                                    {
                                        returnPdfName += ".pdf";                                        
                                    }
                                }
                                else if (sheetName == "")
                                {
                                    returnPdfName = workbookFileName + ".pdf";
                                }
                                else
                                {
                                    returnPdfName = workbookFileName + $"_{sheetName}.pdf";
                                }
                                return returnPdfName;
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Unable to print the following file.\n{filePath}\n\n"+ex.Message, "Error");
                        }
                        
                        if (worker.CancellationPending)
                        {
                            break;
                        }
                    }

                    if (workbookToPrint != null) { workbookToPrint.Close(); }
                    GC.Collect();
                    progressTracker.UpdateStatus("Completed, pending check box");
                    MessageBox.Show("Completed", "Completed");
                }
                finally
                {
                    ThisApplication.ScreenUpdating = true;
                    ThisApplication.DisplayAlerts = true;
                }
                #endregion
            });
        }    
        private void PrintEntireWorkbook(Workbook workbookToPrint, string sheetName, string pdfFileName, string folderPath)
        {
            try
            {
                (bool toProceed, bool toTerminateAll) = CheckPrintPath(folderPath, pdfFileName);
                string pdfFilePath = Path.Combine(folderPath, pdfFileName);

                if (!toProceed || toTerminateAll)
                {
                    throw new Exception("Terminated by user");
                }

                // Print
                workbookToPrint.ExportAsFixedFormat(
                        XlFixedFormatType.xlTypePDF,
                        pdfFilePath,
                        XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: false,
                        FixedFormatExtClassPtr: Type.Missing);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to print wokrbook {workbookToPrint.Name}\n\n" + ex.Message);
            }
        }

        private string[] GetAndCheckExcelFileNames(Range fileNameRange, string extension)
        {
            string[] originalFileNames = GetContentsAsStringArray(fileNameRange, false);
            string[] finalFileNames = new string[originalFileNames.Length];
            List<int> illegalFileIndex = new List<int>();
            for (int i = 0; i < originalFileNames.Length; i++)
            {
                finalFileNames[i] = SanitiseFileName(originalFileNames[i]);

                if (Path.HasExtension(finalFileNames[i]) && Path.GetExtension(finalFileNames[i]) != extension)
                {
                    finalFileNames[i] = Path.ChangeExtension(finalFileNames[i], extension);
                }

                if (originalFileNames[i] != finalFileNames[i])
                {
                    illegalFileIndex.Add(i);
                }
            }

            if (illegalFileIndex.Count > 0)
            {
                #region Get Confirmation
                string msg = "The following file names are invalid and illegal characters will be replaced. Continue with replaced fileName?\n" +
                "Excel text will be updated as well. Formulas will be replaced with text\n";

                foreach (int i in illegalFileIndex)
                {
                    msg += $"\"{originalFileNames[i]}\" updated to \"{finalFileNames}\"\n";
                }

                Confirmation(msg);
                #endregion

                #region Update Excel
                foreach (int i in illegalFileIndex)
                {
                    Range cell = fileNameRange.Cells[i];
                    cell.Value2 = finalFileNames[i];
                }
                #endregion

            }

            return finalFileNames;
        }
        #endregion

        #endregion

        #region Open Files
        private void openFilesInOrder_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Delay
                int sleepDelay = Convert.ToInt32(AttributeTextBoxDic["PdfOpenDelay"].GetDoubleFromTextBox() * 1000);
                #endregion

                #region Get files
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "PDF (*.pdf)|*.pdf|All files (*.*)|*.*";
                DialogResult res = openFileDialog.ShowDialog();
                if (res != DialogResult.OK) { return; }
                string[] filePaths = openFileDialog.FileNames;
                #endregion

                foreach (string filePath in filePaths)
                {
                    System.Diagnostics.Process.Start(filePath);
                    Thread.Sleep(sleepDelay);
                }
                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region Testing Grounds
        private void ProgressBarTest_Click(object sender, EventArgs e)
        {
            //Range selectedRange = ThisApplication.ActiveWindow.RangeSelection;
            //ProgressTracker progressTracker = new ProgressTracker();
            //progressTracker.Show();
            //for (int i = 0; i <= 100; i++)
            //{
            //    progressTracker.UpdateProgress(i);

            //    // Your processing logic here
            //    System.Threading.Thread.Sleep(50); // Simulate some work
            //}
            //progressTracker.Close();

            ProgressHelper.RunWithProgress((worker, progressTrackerLocal) =>
            {
                progressTrackerLocal.UpdateStatus("Executing test process");
                // Simulate work
                for (int i = 0; i <= 100; i++)
                {
                    if (worker.CancellationPending)
                    {
                        break;
                    }
                    if (i == 40)
                    {
                        progressTrackerLocal.UpdateStatus("exceeded 40%");
                    }
                    System.Threading.Thread.Sleep(50); // Simulate work
                    worker.ReportProgress(i);
                }
                //if (worker.CancellationPending) { MessageBox.Show("You Cancelled"); } 
            });
        }



        #endregion

        #region Merge PDF
        private void basicMergePDF_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTrackerLocal) => basicMergePDF_Click_function(worker, progressTrackerLocal));
        }


        private void basicMergePDF_Click_function(BackgroundWorker worker, ProgressTracker progressTrackerLocal)
        {
            #region Check Inputs and Get File Paths from Excel
            progressTrackerLocal.UpdateStatus("Checking Inputs");
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(selectedRange);

            // Check Number of Columns
            int targetColNum = 1;
            if ((endCol - startCol + 1) < targetColNum)
            {
                MessageBox.Show($"Number of columns selected should be {targetColNum}, {endCol - startCol + 1} columns found", "Error");
                return;
            }
            // Read Table 
            List<string> filePaths = new List<string>();
            foreach (Range cell in selectedRange.Cells)
            {
                filePaths.Add(cell.Value2);
            }
            #endregion

            #region Get Output Directory
            string outputPath;
            try
            {
                ((DirectoryTextBox)AttributeTextBoxDic["PdfFolderPath"]).CheckAndGetPath();
                outputPath = MergeFileNameAndDir(dispPdfOutFolder.Text, dispMergeName.Text, ".pdf");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }

            if (!CheckAndDeleteFile(outputPath))
            {
                MessageBox.Show("Process terminated by user", "Terminated");
                return;
            }
            #endregion
            
            #region Merge PDF
            PdfDocument outputDocument = new PdfDocument();
            int numFilesCompleted = 0;
            foreach (string filepath in filePaths)
            {
                progressTrackerLocal.UpdateStatus($"Importing {Path.GetFileName(filepath)}");
                PdfDocument inputDocument = PdfReader.Open(filepath, PdfDocumentOpenMode.Import);
                Dictionary<PdfPage, PdfPage> inputPageTracker = new Dictionary<PdfPage, PdfPage>();// InputPage,OutputPage
                for (int index = 0; index < inputDocument.PageCount; index++)
                {
                    PdfPage inputPage = inputDocument.Pages[index];
                    PdfPage outputPage = outputDocument.AddPage(inputPage);
                    inputPageTracker[inputPage] = outputPage;
                }

                #region Add Bookmarks
                if (createBookmarksCheck.Checked)
                {
                    string bookmarkName = Path.GetFileNameWithoutExtension(filepath);
                    AddBookmarks(inputDocument, outputDocument, bookmarkName, inputPageTracker);
                }
                #endregion

                numFilesCompleted++;
                worker.ReportProgress(ConvertToProgress(numFilesCompleted, filePaths.Count));
                if (worker.CancellationPending)
                {
                    break;
                }
            }
            #endregion

            #region Save PDF
            if (outputDocument.PageCount == 0)
            {
                MessageBox.Show("Final PDF file is empty, no file generated", "Error");
                return;
            }
            outputDocument.Save(outputPath);
            #endregion

            MessageBox.Show("Operation completed", "Success");
        }

        
        #endregion

        #region Advance Merge PDF
        private void advanceMerge_Click(object sender, EventArgs e)
        {
            try
            {
                #region Create table objects           
                ExcelTableCol titleObj = new ExcelTableCol("title", 1);
                ExcelTableCol startPgNumObj = new ExcelTableCol("startPgNum", 2);
                ExcelTableCol endPgNumObj = new ExcelTableCol("endPgNum", 3);
                ExcelTableCol totalPgNumObj = new ExcelTableCol("totalPgnum", 4);
                ExcelTableCol insertNewPageObj = new ExcelTableCol("insertNewPage", 5);
                ExcelTableCol filePathObj = new ExcelTableCol("filePath", 6);

                Dictionary<int, ExcelTableCol> excelTable = new Dictionary<int, ExcelTableCol>();
                excelTable.Add(titleObj.relativeColNum, titleObj);
                excelTable.Add(startPgNumObj.relativeColNum, startPgNumObj);
                excelTable.Add(endPgNumObj.relativeColNum, endPgNumObj);
                excelTable.Add(totalPgNumObj.relativeColNum, totalPgNumObj);
                excelTable.Add(insertNewPageObj.relativeColNum, insertNewPageObj);
                excelTable.Add(filePathObj.relativeColNum, filePathObj);
                #endregion

                #region Check inputs
                Range selectedRange = ThisApplication.ActiveWindow.RangeSelection;
                (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(selectedRange);

                // Check Number of Columns
                CheckRangeSize(selectedRange, 0, 6, "");

                // Check filePath are valid
                bool passCheck = CheckRangeFileExist(selectedRange.Columns[filePathObj.relativeColNum], true, true);
                if (!passCheck)
                {
                    return;
                }
                #endregion

                #region Read excel table
                int colNum = 1;
                foreach (Range colRange in selectedRange.Columns)
                {
                    ExcelTableCol thisColumn = excelTable[colNum];
                    thisColumn.range = colRange;
                    colNum++;
                }
                #endregion

                #region Set array objects 
                string[] headers = titleObj.ConvertRangeToStringArray();
                string[] filePaths = filePathObj.ConvertRangeToStringArray();
                bool[] insertNewPage = insertNewPageObj.ConvertRangeToBoolArray();
                int?[] startPgNum = startPgNumObj.CreateNewIntArray();
                int?[] endPgNum = endPgNumObj.CreateNewIntArray();
                int?[] totalPgNum = totalPgNumObj.CreateNewIntArray();
                #endregion

                #region Get Output Directory
                string outputPath;
                try
                {
                    ((DirectoryTextBox)AttributeTextBoxDic["PdfFolderPath"]).CheckAndGetPath();
                    outputPath = MergeFileNameAndDir(dispPdfOutFolder.Text, dispMergeName.Text, ".pdf");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    return;
                }

                if (!CheckAndDeleteFile(outputPath))
                {
                    MessageBox.Show("Process terminated by user", "Terminated");
                    return;
                }
                #endregion

                #region Get Reference Title File
                string refTitleFilePath = "";
                if (((FileTextBox)AttributeTextBoxDic["RefTitlePageFile"]).textBox.Text != "")
                {
                    try
                    {
                        refTitleFilePath = ((FileTextBox)AttributeTextBoxDic["RefTitlePageFile"]).CheckAndGetValue(true);
                    }
                    catch
                    {
                        return;
                    }
                }

                PdfDocument refTitleFile;
                if (refTitleFilePath != "")
                {
                    refTitleFile = PdfReader.Open(refTitleFilePath, PdfDocumentOpenMode.Import);
                }
                #endregion

                #region Merge PDF
                PdfDocument outputDocument = new PdfDocument();
                int? sectionStartRow = null;
                Dictionary<string, PdfOutline> outlineTracker = new Dictionary<string, PdfOutline>();
                for (int relRowNum = 0; relRowNum < filePaths.Length; relRowNum++) // Loop through excel rows
                {
                    #region Section Title and Page Nums
                    if (headers[relRowNum] != "")
                    {
                        // Set end page number 
                        if (sectionStartRow != null) // skip first occurrence 
                        {
                            endPgNum[(int)sectionStartRow] = outputDocument.PageCount;
                            totalPgNum[(int)sectionStartRow] = endPgNum[(int)sectionStartRow] - startPgNum[(int)sectionStartRow] + 1;
                        }
                        // Set start page number
                        sectionStartRow = relRowNum;
                        startPgNum[relRowNum] = outputDocument.PageCount + 1;

                        // Insert New Page if required
                        if (insertNewPage[relRowNum])
                        {
                            PdfPage page;
                            if (refTitleFilePath != "")
                            {
                                refTitleFile = PdfReader.Open(refTitleFilePath, PdfDocumentOpenMode.Import);
                                page = outputDocument.AddPage(refTitleFile.Pages[0]);
                            }
                            else
                            {
                                page = outputDocument.AddPage();
                            }
                            //PdfPage page = outputDocument.AddPage();
                            InsertHeaderPage(page, headers[relRowNum]);
                            sectionStartRow = relRowNum;
                        }
                    }
                    #endregion

                    #region Append Input File
                    if (filePaths[relRowNum] == "")
                    {
                        // Skip empty filepaths
                        continue;
                    }

                    PdfDocument inputDocument = PdfReader.Open(filePaths[relRowNum], PdfDocumentOpenMode.Import);
                    for (int inputPageNum = 0; inputPageNum < inputDocument.PageCount; inputPageNum++)
                    {
                        PdfPage page = inputDocument.Pages[inputPageNum];
                        outputDocument.AddPage(page);
                    }
                    #endregion

                    #region Add Bookmark
                    if (headers[relRowNum] != "")
                    {
                        string header = headers[relRowNum];
                        string[] parts = header.Split('\n');
                        PdfPage bookmarkPage = outputDocument.Pages[(int)startPgNum[relRowNum] - 1];

                        if (parts.Count() == 2)
                        {
                            string parentBookmarkName = parts[0];
                            string childBookmarkName = parts[1];
                            PdfOutline parentBookmark = null;
                            if (outlineTracker.ContainsKey(parentBookmarkName))
                            { 
                                parentBookmark = outlineTracker[parentBookmarkName];
                            }
                            else
                            {
                                parentBookmark = outputDocument.Outlines.Add(parentBookmarkName, bookmarkPage);
                                outlineTracker.Add(parentBookmarkName, parentBookmark);
                            }
                            if (childBookmarkName.Length > 0) { parentBookmark.Outlines.Add(childBookmarkName, bookmarkPage); }
                        }
                        else
                        {
                            PdfOutline bookmark = outputDocument.Outlines.Add(header, bookmarkPage);
                            outlineTracker.Add(header, bookmark);
                        }
                    }
                    #endregion
                }
                #region Write Page Numbers
                // Calculate final occurrence of header
                if (sectionStartRow != null) // skip if no headers
                {
                    endPgNum[(int)sectionStartRow] = outputDocument.PageCount;
                    totalPgNum[(int)sectionStartRow] = endPgNum[(int)sectionStartRow] - startPgNum[(int)sectionStartRow] + 1;
                }
                startPgNumObj.WriteIntToExcel();
                endPgNumObj.WriteIntToExcel();
                totalPgNumObj.WriteIntToExcel();
                #endregion
                #endregion

                #region Save PDF
                if (outputDocument.PageCount == 0)
                {
                    MessageBox.Show("Final PDF file is empty, no file generated", "Error");
                    return;
                }
                outputDocument.Save(outputPath);
                #endregion

                MessageBox.Show("Operation completed", "Success");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            
        }

        private void InsertHeaderPage(PdfPage page, string headerText)
        {
            XGraphics gfx = XGraphics.FromPdfPage(page);
            string fontName = "Arial";
            int fontSize = Convert.ToInt32(dispTitleFontSize.Text);
            XFont fontType = new XFont(fontName, fontSize);

            XRect rect = new XRect(page.Width.Value/6, page.Height.Value/2-50, page.Width.Value/6*4, page.Height.Value-60);
            string textContents = headerText;
            XTextFormatter tf = new XTextFormatter(gfx);
            tf.LayoutRectangle = rect;
            tf.Alignment = XParagraphAlignment.Center;
            tf.DrawString(textContents, fontType, XBrushes.Black, rect, XStringFormats.TopLeft);
        }
        #endregion
        
        #region Add Bookmars
        void AddBookmarks(PdfDocument inputDocument, PdfDocument outputDocument, string parentBookmarkName, Dictionary<PdfPage, PdfPage> inputPageTracker, PdfOutline grandParentOutline = null)
        {
            #region Add Base Bookmark
            PdfPage firstInsertPage = outputDocument.Pages[outputDocument.PageCount - inputDocument.PageCount];
            PdfOutline parentOutline;
            if (grandParentOutline == null)
            {
                parentOutline = outputDocument.Outlines.Add(parentBookmarkName, firstInsertPage);
            }
            else
            {
                parentOutline = grandParentOutline.Outlines.Add(parentBookmarkName, firstInsertPage);
            }

            #endregion
            try
            {
                //inputDocument.Outlines is giving me issues when it is empty, I don't know why and i can't seem detect when it is empty (simply accessing inputDocument.Outlines is an error).
                PdfOutlineCollection thisCollection = inputDocument.Outlines;
            }
            catch { return; }

            AddNestedBookmarks(inputDocument, outputDocument, inputPageTracker, parentOutline, inputDocument.Outlines);

        }
        void AddNestedBookmarks(PdfDocument inputDocument, PdfDocument outputDocument, Dictionary<PdfPage, PdfPage> inputPageTracker, PdfOutline parentOutline, PdfOutlineCollection inputOutlineCollection)
        {
            //try
            //{
            //    if (inputOutlineCollection.Count == 0) { return; }
            //    foreach (PdfOutline inputOutline in inputOutlineCollection)
            //    {
            //        // Create a new bookmark in the output document
            //        PdfPage inputPage = inputOutline.DestinationPage;
            //        PdfPage outputPage = inputPageTracker[inputPage];
            //        PdfOutline newOutline = parentOutline.Outlines.Add(inputOutline.Title, outputPage);

            //        try
            //        {
            //            //inputDocument.Outlines is giving me issues when it is empty, I don't know why and i can't seem detect when it is empty (simply accessing inputDocument.Outlines is an error).
            //            PdfOutlineCollection thisCollection = inputDocument.Outlines;
            //        }
            //        catch { continue; }
            //        AddNestedBookmarks(inputDocument, outputDocument, inputPageTracker, newOutline, inputOutline.Outlines);
            //    }
            //}

            //catch (Exception ex) { }

            if (inputOutlineCollection.Count == 0) { return; }
            foreach (PdfOutline inputOutline in inputOutlineCollection)
            {
                // Create a new bookmark in the output document
                PdfPage inputPage = inputOutline.DestinationPage;
                PdfPage outputPage = inputPageTracker[inputPage];
                PdfOutline newOutline = parentOutline.Outlines.Add(inputOutline.Title, outputPage);

                try
                {
                    //inputDocument.Outlines is giving me issues when it is empty, I don't know why and i can't seem detect when it is empty (simply accessing inputDocument.Outlines is an error).
                    PdfOutlineCollection thisCollection = inputDocument.Outlines;
                }
                catch { continue; }
                AddNestedBookmarks(inputDocument, outputDocument, inputPageTracker, newOutline, inputOutline.Outlines);
            }
        }
        #endregion 
        
        #region Generate Section Title PDF
        private void generateSections_Click(object sender, EventArgs e)
        {
            #region Create table objects           
            int colNum = 1;
            ExcelTableCol fileNameObj = new ExcelTableCol("fileName", colNum);
            colNum++;
            ExcelTableCol titleTextObj = new ExcelTableCol("titleText", 2);
            colNum++;

            Dictionary<int, ExcelTableCol> excelTable = new Dictionary<int, ExcelTableCol>();
            excelTable.Add(fileNameObj.relativeColNum, fileNameObj);
            excelTable.Add(titleTextObj.relativeColNum, titleTextObj);
            #endregion

            #region Check inputs
            Range selectedRange = ThisApplication.ActiveWindow.RangeSelection;
            (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(selectedRange);

            // Check Number of Columns
            int targetColNum = excelTable.Keys.Count;
            if ((endCol - startCol + 1) != targetColNum)
            {
                MessageBox.Show($"Number of columns selected should be {targetColNum}, {endCol - startCol + 1} columns found", "Error");
                return;
            }
            #endregion

            #region Read excel table
            colNum = 1;
            foreach (Range colRange in selectedRange.Columns)
            {
                ExcelTableCol thisColumn = excelTable[colNum];
                thisColumn.range = colRange;
                colNum++;
            }

            // Check if range provided is empty
            if (!CheckRangeIsFilled(fileNameObj.range, true))
            {
                return;
            }
            if (!CheckRangeIsFilled(titleTextObj.range, true))
            {
                return;
            }
            #endregion

            #region Set array objects 
            string[] fileName = fileNameObj.ConvertRangeToStringArray();
            string[] titleText = titleTextObj.ConvertRangeToStringArray();
            #endregion

            #region Get Output Directory and Check Output Files
            string outputFolder;
            try
            {
                outputFolder = ((DirectoryTextBox)AttributeTextBoxDic["PdfFolderPath"]).CheckAndGetPath();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }

            string[] outputPathArray = new string[titleText.Length];
            for (int index = 0; index <  fileName.Length; index++)
            {
                string outputPath = MergeFileNameAndDir(dispPdfOutFolder.Text, fileName[index], ".pdf");
                if (!CheckAndDeleteFile(outputPath))
                {
                    MessageBox.Show("Process terminated by user", "Terminated");
                    return;
                }
                outputPathArray[index] = outputPath;
            }
            #endregion

            #region Get Reference Title File
            string refTitleFilePath = "";
            if (((FileTextBox)AttributeTextBoxDic["RefTitlePageFile"]).textBox.Text != "")
            {
                try
                {
                    refTitleFilePath = ((FileTextBox)AttributeTextBoxDic["RefTitlePageFile"]).CheckAndGetValue(true);
                }
                catch
                {
                    return;
                }
            }
            
            PdfDocument refTitleFile;
            if (refTitleFilePath != "")
            {
                refTitleFile = PdfReader.Open(refTitleFilePath, PdfDocumentOpenMode.Import);
            }
            #endregion

            #region Create Files
            for (int index = 0; index < fileName.Length; index++)
            {
                #region Create Document
                PdfDocument outputDocument = new PdfDocument();
                PdfPage page;
                if (refTitleFilePath != "")
                {
                    refTitleFile = PdfReader.Open(refTitleFilePath, PdfDocumentOpenMode.Import);
                    page = outputDocument.AddPage(refTitleFile.Pages[0]);
                }
                else
                {
                    page = outputDocument.AddPage();
                }
                InsertHeaderPage(page, titleText[index]);
                #endregion

                #region Save File
                if (outputDocument.PageCount == 0)
                {
                    MessageBox.Show("Final PDF file is empty, no file generated", "Error");
                    return;
                }
                outputDocument.Save(outputPathArray[index]);
                #endregion
            }
            #endregion
            MessageBox.Show("Operation completed", "Success");
        }
        #endregion

        #region Add Page Number
        private void addPageNum_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            if (fileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string[] inputPaths = fileDialog.FileNames;

            ProgressHelper.RunWithProgress((worker, progressTrackerLocal) => RunFunction(worker, progressTrackerLocal));

            void RunFunction(BackgroundWorker worker, ProgressTracker progressTrackerLocal)
            {
                #region Check Inputs
                {
                    bool success = double.TryParse(dispOffsetX.Text, out double checkdouble);
                    if (!success)
                    {
                        MessageBox.Show($"Unable to convert {dispOffsetX.Text} to number for Page Number Offset X");
                        return;
                    }

                    success = double.TryParse(dispOffsetY.Text, out checkdouble);
                    if (!success)
                    {
                        MessageBox.Show($"Unable to convert {dispOffsetX.Text} to number for Page Number Offset X");
                        return;
                    }
                }
                #endregion

                #region Add Page Num
                try
                {
                    foreach (string inputPath in inputPaths)
                    {
                        AddPageNumToOne(inputPath, worker, progressTrackerLocal);
                        if (worker.CancellationPending)
                        {
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unable to complete operation" +
                        $"\n\nError message:" +
                        $"\n{ex.Message}", "Error");
                    return;
                }
                #endregion

                #region Open Files and Show Completion
                if (!worker.CancellationPending)
                {
                    if (checkOpenOutput.Checked)
                    {
                        foreach (string inputPath in inputPaths)
                        {
                            string baseDirectory = Path.GetDirectoryName(inputPath);
                            string inputFileName = Path.GetFileNameWithoutExtension(inputPath);
                            string outputFileName = inputFileName + $"{dispAppendName.Text}.pdf";
                            string outputPath = Path.Combine(baseDirectory, outputFileName);
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
                        }
                    }
                    MessageBox.Show("Operation completed", "Success");
                }
                #endregion
            }
        }

        private void AddPageNumToOne(string inputPath, BackgroundWorker worker, ProgressTracker progressTrackerLocal)
        {
            #region Check if file exist
            progressTrackerLocal.UpdateStatus($"Checking inputs for {Path.GetFileName(inputPath)}");
            if (!File.Exists(inputPath))
            {
                throw new Exception($"File does not exist at {inputPath}");
            }
            #endregion

            #region Get Input and Output Paths
            string baseDirectory = Path.GetDirectoryName(inputPath);
            string inputFileName = Path.GetFileNameWithoutExtension(inputPath);
            string outputFileName = inputFileName + $"{dispAppendName.Text}.pdf";
            string outputPath = Path.Combine(baseDirectory, outputFileName);

            if (!CheckAndDeleteFile(outputPath))
            {
                MessageBox.Show("Process terminated by user", "Terminated");
                return;
            }
            #endregion

            #region Get Page Number Parameters and Check Appended File Name
            // First Page Number
            if (dispFirstPageNum.Text == "")
            {
                dispFirstPageNum.Text = "1";
            }

            int firstPageNum;
            try
            {
                firstPageNum = Convert.ToInt32(dispFirstPageNum.Text);
            }
            catch
            {
                MessageBox.Show($"Error, unable to convert \"{dispFirstPageNum.Text}\" to integer for First Page Number", "Error");
                return;
            }

            if (dispSkipPage.Text == "")
            {
                dispSkipPage.Text = "0";
            }
            // Ignore First N Pages
            int skipPages;
            try
            {
                skipPages = Convert.ToInt32(dispSkipPage.Text);
            }
            catch
            {
                MessageBox.Show($"Error, unable to convert \"{dispSkipPage.Text}\" to integer for Ignore First N Pages", "Error");
                return;
            }
            // Append File Name
            if (dispAppendName.Text == "")
            {
                dispAppendName.Text = "_withPgNum";
            }
            #endregion

            #region Add Page Number
            progressTrackerLocal.UpdateStatus($"Adding page number to {Path.GetFileName(inputPath)}");
            PdfDocument inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Modify);
            
            //PdfDocument outputDocument = new PdfDocument();
            //foreach (PdfPage page in inputDocument.Pages)
            //{
            //    outputDocument.AddPage(page);
            //}

            int currentPageNum = firstPageNum;
            //for (int pageNum = skipPages; pageNum < outputDocument.PageCount; pageNum++)
            for (int pageNum = skipPages; pageNum < inputDocument.PageCount; pageNum++)
            {
                #region User Params
                int fontSize = Convert.ToInt32(dispFontSize.Text);
                double xOffset = double.Parse(dispOffsetX.Text);
                double yOffset = double.Parse(dispOffsetY.Text);
                #endregion

                //PdfPage page = outputDocument.Pages[pageNum];
                PdfPage page = inputDocument.Pages[pageNum];
                XGraphics gfx = XGraphics.FromPdfPage(page);
                string fontName = "Arial";

                XFont fontType = new XFont(fontName, fontSize);

                int rotation = page.Rotate;
                XPoint bottomLeftPoint;
                switch (rotation)
                {
                    case 90:
                        gfx.RotateTransform(-90);
                        bottomLeftPoint = new XPoint(-xOffset, page.Width.Value - yOffset);
                        break;
                    case 180:
                        gfx.RotateTransform(-180);
                        bottomLeftPoint = new XPoint(-xOffset, -yOffset);
                        break;
                    case 270:
                        gfx.RotateTransform(-270);
                        bottomLeftPoint = new XPoint(page.Height.Value - xOffset, -yOffset);
                        break;
                    default:
                        bottomLeftPoint = new XPoint(page.Width.Value - xOffset, page.Height.Value - yOffset);
                        break;
                }

                string textContents = $"{currentPageNum}";
                XSize textSize = gfx.MeasureString(textContents, fontType);
                XPoint topRightPoint = new XPoint(bottomLeftPoint.X - textSize.Width, bottomLeftPoint.Y - (textSize.Height));
                XRect rect = new XRect(topRightPoint, bottomLeftPoint);
                gfx.DrawRectangle(XBrushes.White, rect);
                gfx.DrawString(textContents, fontType, XBrushes.Black, rect, XStringFormats.BottomRight);


                #region Report Progress
                worker.ReportProgress(ConvertToProgress(pageNum, inputDocument.PageCount));
                if (worker.CancellationPending)
                {
                    return;
                }
                #endregion

                currentPageNum++;
            }
            #endregion
            inputDocument.Save(outputPath);
        }
        #endregion

        #region Tests
        FontDialog fontDialog = new FontDialog();
        private void initFont()
        {
            fontDialog.ShowColor = true;
        }

        private void titleFont_Click(object sender, EventArgs e)
        {
            fontDialog.ShowDialog();
        }

        #endregion

        #region Compare Folders
        private void union_Click(object sender, EventArgs e)
        {
            CompareFolders("union");
        }

        private void intersectFiles_Click(object sender, EventArgs e)
        {
            CompareFolders("intersect");
        }

        private void subtractFiles_Click(object sender, EventArgs e)
        {
            CompareFolders("subtract");
        }

        private void removeIntersectFiles_Click(object sender, EventArgs e)
        {
            CompareFolders("removeIntersect");
        }

        private void reverseSubtractFiles_Click(object sender, EventArgs e)
        {
            CompareFolders("reverseSubtract");
        }
        
        private void CompareFolders(string comparisonType)
        {
            try
            {
                #region Get Paths
                string folderPath1 = ((DirectoryTextBox)AttributeTextBoxDic["FolderPath1"]).CheckAndGetPath();
                HashSet<string> files1 = GetFileNamesOnly(folderPath1);

                string folderPath2 = ((DirectoryTextBox)AttributeTextBoxDic["FolderPath2"]).CheckAndGetPath();
                HashSet<string> files2 = GetFileNamesOnly(folderPath2);
                #endregion

                #region Compare
                string[] resultant;
                switch (comparisonType)
                {
                    case "union":
                        {
                            resultant = files1.Union(files2).OrderBy(s => s).ToArray();
                            break;
                        }
                    case "intersect":
                        {
                            resultant = files1.Intersect(files2).OrderBy(s => s).ToArray();
                            break;
                        }
                    case "subtract":
                        {
                            resultant = files1.Except(files2).OrderBy(s => s).ToArray();
                            break;
                        }
                    case "removeIntersect":
                        {
                            resultant = files1.Except(files2).Union(files2.Except(files1)).OrderBy(s => s).ToArray();
                            break;
                        }
                    case "reverseSubtract":
                        {
                            resultant = files2.Except(files1).OrderBy(s => s).ToArray();
                            break;
                        }
                    default:
                        throw new ArgumentException($"Comparison Type {comparisonType} not valid.");
                }
                #endregion

                WriteToExcelSelectionAsRow(0, 0, true, resultant);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        
        private HashSet<string> GetFileNamesOnly(string folderPath)
        {

            string extension = "";
            if (specifyExtensionCheck.Checked)
            {
                extension = AttributeTextBoxDic["ExtensionTypeComparison"].textBox.Text;
            }

            List<string> files1 = new List<string>();
            getFiles(folderPath, ref files1, searchSubFoldersCheck.Checked, extension);
            HashSet<string> fileNames = new HashSet<string>();
            foreach (string file in files1)
            {
                string fileName = Path.GetFileName(file);
                if (!fileNames.Contains(file)) { fileNames.Add(fileName); }
            }
            return fileNames;
        }


        #endregion

        #region Renumber Sheets
        private void getFileInfo_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get files
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "PDF (*.pdf)|*.pdf";
                DialogResult res = openFileDialog.ShowDialog();
                if (res != DialogResult.OK) { return; }
                string[] filePaths = openFileDialog.FileNames;
                #endregion

                #region Number Files
                string[] number = new string[filePaths.Length];
                string[] fileNames = new string[filePaths.Length];
                for (int i = 0; i < filePaths.Length; i++)
                {
                    fileNames[i] = Path.GetFileNameWithoutExtension(filePaths[i]);
                    number[i] = (i + 1).ToString("000");
                }
                #endregion

                #region Update Print Range Format
                {
                    Range currentSelection = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                    Range startRange = currentSelection.Offset[0, 2];
                    Range endRange = startRange.Offset[fileNames.Length - 1, 0];
                    Range numberRange = currentSelection.Worksheet.Range[startRange, endRange];
                    numberRange.NumberFormat = "@";
                }
                #endregion
                
                WriteToExcelRangeAsCol(null, 0, 0, true, filePaths, fileNames, number);
                dispTotalDwgNum.Text = filePaths.Length.ToString();
            }
            catch (Exception ex) { MessageBox.Show($"Error:{ex.Message}"); }
        }

        private void editFilesSheetNum_Click(object sender, EventArgs e)
        {
            try
            {
                #region Read input data
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                string[] filePaths = GetContentsAsStringArray(selectedRange.Columns[1], false);
                string[] fileNames = GetContentsAsStringArray(selectedRange.Columns[2], false);
                string[] fileNum = GetContentsAsStringArray(selectedRange.Columns[3], false);

                int fontSize = Convert.ToInt32(dispFontSizeSheetNum.Text);
                double xSheetNum = double.Parse(dispThisSheetX.Text) * 72 / 25.4;
                double ySheetNum = double.Parse(dispThisSheetY.Text) * 72 / 25.4;
                double xTotalSheetNum = double.Parse(dispTotalSheetX.Text) * 72 / 25.4;
                double yTotalSheetNum = double.Parse(dispTotalSheetY.Text) * 72 / 25.4;

                #endregion

                #region Create and Empty Destination
                string workbookPath = Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
                string printPath = Path.Combine(workbookPath, "Updated Dwg");
                CreateDestinationFolder(printPath);
                bool overwriteExisting = ClearFolder(printPath);
                #endregion

                #region Copy files over
                string[] finalFilePaths = new string[filePaths.Length];
                for (int i = 0; i < filePaths.Length; i++)
                {
                    string filePath = filePaths[i];
                    string fileName = Path.GetFileName(filePath);

                    string finalFileName;
                    if (renameFilesCheck.Checked)
                    {
                        finalFileName = fileNum[i] + "_" + fileName;
                    }
                    else
                    {
                        finalFileName = fileName;
                    }

                    string destinationPath = Path.Combine(printPath, finalFileName);

                    File.Copy(filePath, destinationPath, overwriteExisting);
                    finalFilePaths[i] = destinationPath;
                }
                #endregion

                #region Add Sheet number
                if (addSheetNumberCheck.Checked)
                {
                    for (int i =0; i < finalFilePaths.Length; i++)
                    {
                        AddSheetNumberToOne(finalFilePaths[i], fileNum[i], dispTotalDwgNum.Text,
                                            fontSize, xSheetNum, ySheetNum, xTotalSheetNum, yTotalSheetNum);
                    } 
                }
                #endregion
                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show($"Error:{ex.Message}"); }
        }

        private void AddSheetNumberToOne(string filePath, string sheetNum, string totalSheetNum,
                            int fontSize, double xSheetNum, double ySheetNum, double xTotalSheetNum, double yTotalSheetNum)
        {
            #region Check if file exist
            if (!File.Exists(filePath))
            {
                throw new Exception($"File does not exist at {filePath}");
            }
            #endregion

            #region Open File
            PdfDocument inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Modify);
            PdfPage page = inputDocument.Pages[0];
            XGraphics gfx = XGraphics.FromPdfPage(page);
            string fontName = "Arial";
            XFont fontType = new XFont(fontName, fontSize);

            // Testing values
            double width = page.Width.Value;
            double height = page.Height.Value;
            // End test

            int rotation = page.Rotate;
            XPoint bottomLeftPoint;
            switch (rotation)
            {
                case 90:
                    gfx.RotateTransform(-90);
                    bottomLeftPoint = new XPoint(-xSheetNum, page.Width.Value - ySheetNum);
                    break;
                case 180:
                    gfx.RotateTransform(-180);
                    bottomLeftPoint = new XPoint(-xSheetNum, -ySheetNum);
                    break;
                case 270:
                    gfx.RotateTransform(-270);
                    bottomLeftPoint = new XPoint(page.Height.Value - xSheetNum, -ySheetNum);
                    break;
                default:
                    bottomLeftPoint = new XPoint(xSheetNum, ySheetNum);
                    break;
            }

            string textContents = sheetNum;
            XSize textSize = gfx.MeasureString(textContents, fontType);
            XPoint topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - (textSize.Height));
            XRect rect = new XRect(topRightPoint, bottomLeftPoint);
            //gfx.DrawRectangle(XBrushes.White, rect);
            gfx.DrawRectangle(XBrushes.LightBlue, rect);
            //gfx.DrawString(textContents, fontType, XBrushes.Black, rect, XStringFormats.BottomRight);
            gfx.DrawString(textContents, fontType, XBrushes.Blue, rect, XStringFormats.BottomRight);


            bottomLeftPoint = new XPoint(xTotalSheetNum, yTotalSheetNum);
            textContents = totalSheetNum;
            textSize = gfx.MeasureString(textContents, fontType);
            topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - textSize.Height);
            rect = new XRect(topRightPoint, bottomLeftPoint);
            gfx.DrawRectangle(XBrushes.LightPink, rect);
            #endregion

            inputDocument.Save(filePath);
        }

        private void AddPageNumToOneRef(string inputPath, BackgroundWorker worker, ProgressTracker progressTrackerLocal)
        {
            #region Check if file exist
            progressTrackerLocal.UpdateStatus($"Checking inputs for {Path.GetFileName(inputPath)}");
            if (!File.Exists(inputPath))
            {
                throw new Exception($"File does not exist at {inputPath}");
            }
            #endregion

            #region Get Input and Output Paths
            string baseDirectory = Path.GetDirectoryName(inputPath);
            string inputFileName = Path.GetFileNameWithoutExtension(inputPath);
            string outputFileName = inputFileName + $"{dispAppendName.Text}.pdf";
            string outputPath = Path.Combine(baseDirectory, outputFileName);

            if (!CheckAndDeleteFile(outputPath))
            {
                MessageBox.Show("Process terminated by user", "Terminated");
                return;
            }
            #endregion

            #region Get Page Number Parameters and Check Appended File Name
            // First Page Number
            if (dispFirstPageNum.Text == "")
            {
                dispFirstPageNum.Text = "1";
            }

            int firstPageNum;
            try
            {
                firstPageNum = Convert.ToInt32(dispFirstPageNum.Text);
            }
            catch
            {
                MessageBox.Show($"Error, unable to convert \"{dispFirstPageNum.Text}\" to integer for First Page Number", "Error");
                return;
            }

            if (dispSkipPage.Text == "")
            {
                dispSkipPage.Text = "0";
            }
            // Ignore First N Pages
            int skipPages;
            try
            {
                skipPages = Convert.ToInt32(dispSkipPage.Text);
            }
            catch
            {
                MessageBox.Show($"Error, unable to convert \"{dispSkipPage.Text}\" to integer for Ignore First N Pages", "Error");
                return;
            }
            // Append File Name
            if (dispAppendName.Text == "")
            {
                dispAppendName.Text = "_withPgNum";
            }
            #endregion

            #region Add Page Number
            progressTrackerLocal.UpdateStatus($"Adding page number to {Path.GetFileName(inputPath)}");
            PdfDocument inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Modify);

            //PdfDocument outputDocument = new PdfDocument();
            //foreach (PdfPage page in inputDocument.Pages)
            //{
            //    outputDocument.AddPage(page);
            //}

            int currentPageNum = firstPageNum;
            //for (int pageNum = skipPages; pageNum < outputDocument.PageCount; pageNum++)
            for (int pageNum = skipPages; pageNum < inputDocument.PageCount; pageNum++)
            {
                #region User Params
                int fontSize = Convert.ToInt32(dispFontSize.Text);
                double xOffset = double.Parse(dispOffsetX.Text);
                double yOffset = double.Parse(dispOffsetY.Text);
                #endregion

                //PdfPage page = outputDocument.Pages[pageNum];
                PdfPage page = inputDocument.Pages[pageNum];
                XGraphics gfx = XGraphics.FromPdfPage(page);
                string fontName = "Arial";

                XFont fontType = new XFont(fontName, fontSize);

                int rotation = page.Rotate;
                XPoint bottomLeftPoint;
                switch (rotation)
                {
                    case 90:
                        gfx.RotateTransform(-90);
                        bottomLeftPoint = new XPoint(-xOffset, page.Width.Value - yOffset);
                        break;
                    case 180:
                        gfx.RotateTransform(-180);
                        bottomLeftPoint = new XPoint(-xOffset, -yOffset);
                        break;
                    case 270:
                        gfx.RotateTransform(-270);
                        bottomLeftPoint = new XPoint(page.Height.Value - xOffset, -yOffset);
                        break;
                    default:
                        bottomLeftPoint = new XPoint(page.Width.Value - xOffset, page.Height.Value - yOffset);
                        break;
                }

                string textContents = $"{currentPageNum}";
                XSize textSize = gfx.MeasureString(textContents, fontType);
                XPoint topRightPoint = new XPoint(bottomLeftPoint.X - textSize.Width, bottomLeftPoint.Y - (textSize.Height));
                XRect rect = new XRect(topRightPoint, bottomLeftPoint);
                gfx.DrawRectangle(XBrushes.White, rect);
                gfx.DrawString(textContents, fontType, XBrushes.Black, rect, XStringFormats.BottomRight);


                #region Report Progress
                worker.ReportProgress(ConvertToProgress(pageNum, inputDocument.PageCount));
                if (worker.CancellationPending)
                {
                    return;
                }
                #endregion

                currentPageNum++;
            }
            #endregion
            inputDocument.Save(outputPath);
        }

        private bool ClearFolder(string printPath)
        {
            if (!Directory.Exists(printPath)) { throw new Exception($"Path does not exist {printPath}"); }

            string[] files = Directory.GetFiles(printPath);
            if (files.Length == 0) { return false; }

            DialogResult res = MessageBox.Show($"Delete all files in folder {printPath}?", "Warning", MessageBoxButtons.YesNoCancel);
            if (res == DialogResult.Cancel) { throw new Exception($"Terminated by user"); }
            else if (res == DialogResult.Yes)
            {
                foreach (string file in files)
                {
                    File.Delete(file);
                }
                return false;
            }
            else 
            {
                DialogResult res2 = MessageBox.Show($"Overwrite existing files?", "Warning", MessageBoxButtons.YesNoCancel);
                if (res2 == DialogResult.Cancel) { throw new Exception($"Terminated by user"); }
                else if (res2 == DialogResult.Yes)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        #endregion


    }
}

