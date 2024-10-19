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

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class PrintToolsPane : UserControl
    {
        #region Initialisers
        Workbook ThisWorkBook;
        Microsoft.Office.Interop.Excel.Application ThisApplication;
        DocumentProperties AllCustProps;
        Dictionary<string, AttributeTextBox> AttributeTextBoxDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> CustomAttributeDic = new Dictionary<string, CustomAttribute>();

        public PrintToolsPane()
        {
            InitializeComponent();
            ThisApplication = Globals.ThisAddIn.Application;
            ThisWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            AllCustProps = ThisWorkBook.CustomDocumentProperties;
            CreateAttributes();
            AddToolTips();
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

            #region PDF
            DirectoryTextBox PdfFolderPath = new DirectoryTextBox("PdfFolderPath", dispPdfOutFolder, setPdfOutFolder);
            PdfFolderPath.AddOpenButton(openPdfOutFolder);
            AttributeTextBoxDic.Add("PdfFolderPath", PdfFolderPath);

            AttributeTextBox MergeName = new AttributeTextBox("MergeName", dispMergeName, true);
            MergeName.type = "filename";
            AttributeTextBoxDic.Add("MergeName", MergeName);

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

            #region Merge PDF
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
            toolTip1.SetToolTip(advancedMerge,
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
                        MessageBox.Show("Unable to overwrite file, please check if file is open.", "Failed to overwrite");
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
                CommonUtilities.WriteToExcel(0, 0, true, names.ToArray());
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
            try { CheckRangeSize(selectedRange, 0, 2); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }
            string[] sheetNames = GetContentsAsStringArray(selectedRange.Columns[1].Cells, false);
            string[] filePaths = GetContentsAsStringArray(selectedRange.Columns[2].Cells, false);
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
                        string pdfFileName = Path.GetFileNameWithoutExtension(filePath);

                        #endregion

                        #region Print Sheets
                        if (sheetName == "") // Print all visible workbook
                        {
                            pdfFileName += ".pdf";
                            PrintEntireWorkbook(workbookToPrint, sheetName, pdfFileName, folderPath);
                        }
                        else // Print single sheet
                        {
                            pdfFileName += $"_{sheetName}.pdf";
                            GetAndPrintSingleSheet(workbookToPrint, sheetName, pdfFileName, folderPath);
                        }
                        #endregion

                        numPrinted++;
                        worker.ReportProgress(ConvertToProgress(rowNum+1, filePaths.Length));
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

        private void insertPrintWorkbookHeader_Click(object sender, EventArgs e)
        {
            List<string> headers = new List<string> { "Sheet Name - leave blank to print all", "File Path" };
            InsertHeadersAtSelection(headers, "cols");
        }
        #endregion

        #endregion

        #region Directory Management
        //private void importFilePath_Click(object sender, EventArgs e)
        //{
        //    #region Check Directory
        //    if (dispDirectory.Text == "")
        //    {
        //        MessageBox.Show("Please provide folderpath", "Error");
        //        return;
        //    }
        //    else if (!Directory.Exists(dispDirectory.Text))
        //    {
        //        MessageBox.Show($"Invalid folder path:\n\n{dispDirectory.Text}", "Error");
        //        return;
        //    }
        //    #endregion

        //    #region Get Parameters
        //    string directoryPath = dispDirectory.Text;
        //    Workbook activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Worksheet activeSheet = activeBook.ActiveSheet;
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
        //    #endregion

        //    #region Call method to get files and folders
        //    //List<string> directories = new List<string>();
        //    List<string> files = new List<string>();
        //    getFilesAndDirectories(directoryPath, ref files, checkNestedFolders.Checked);
        //    #endregion

        //    #region Print results
        //    // Print files array
        //    string[] folder_name = new string[files.Count()];
        //    string[] file_name = new string[files.Count()];
        //    string[] full_path = new string[files.Count()];
        //    int i = 0;
        //    foreach (string file in files)
        //    {
        //        full_path[i] = file;
        //        if (addExtensionCheck.Checked) { file_name[i] = Path.GetFileName(file); }
        //        else { file_name[i] = Path.GetFileNameWithoutExtension(file); }
        //        folder_name[i] = Path.GetFileName(Path.GetDirectoryName(file));
        //        i++;
        //    }
        //    try
        //    {
        //        WriteToExcel(0, 0, true, full_path, folder_name, file_name);
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message == "Nothing found to print")
        //        {
        //            MessageBox.Show("No results found", "Error");
        //        }
        //        else
        //        {
        //            MessageBox.Show($"Error encountered\n\n{ex.Message}", "Error");
        //        }
        //    }
        //    #endregion
        //}

        //private void importFolderPath_Click(object sender, EventArgs e)
        //{
        //    #region Check Directory
        //    if (dispDirectory.Text == "")
        //    {
        //        MessageBox.Show("Please provide folderpath");
        //        return;
        //    }
        //    else if (!Directory.Exists(dispDirectory.Text))
        //    {
        //        MessageBox.Show($"Invalid folder path:\n\n{dispDirectory.Text}");
        //        return;
        //    }
        //    #endregion

        //    #region Get Parameters
        //    string directoryPath = dispDirectory.Text;
        //    Workbook activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Worksheet activeSheet = activeBook.ActiveSheet;
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
        //    #endregion

        //    #region Call method to get files and folders
        //    List<string> directories = new List<string>();
        //    getFolderDirectories(directoryPath, ref directories, checkNestedFolders.Checked);
        //    #endregion

        //    #region Print results
        //    // Print files array
        //    string[] folder_name = new string[directories.Count()];
        //    string[] full_path = new string[directories.Count()];
        //    int i = 0;
        //    foreach (string folder in directories)
        //    {
        //        full_path[i] = folder;
        //        //folder_name[i] = Path.GetFileName(Path.GetDirectoryName(file));
        //        folder_name[i] = Path.GetFileName(folder);
        //        i++;
        //    }
        //    try
        //    {
        //        WriteToExcel(0, 0, true, full_path, folder_name);
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message == "Nothing found to print")
        //        {
        //            MessageBox.Show("No results found");
        //        }
        //        else
        //        {
        //            MessageBox.Show($"Error encountered\n\n{ex.Message}");
        //        }
        //    }
        //    #endregion
        //}

        //private void importSpecificFile_Click(object sender, EventArgs e)
        //{
        //    #region Check Directory and Inputs
        //    if (dispDirectory.Text == "")
        //    {
        //        MessageBox.Show("Please provide folderpath", "Error");
        //        return;
        //    }
        //    else if (!Directory.Exists(dispDirectory.Text))
        //    {
        //        MessageBox.Show($"Invalid folder path:\n\n{dispDirectory.Text}", "Error");
        //        return;
        //    }

        //    if (dispExtension.Text == "")
        //    {
        //        MessageBox.Show("Please provide extension type", "Error");
        //        return;
        //    }
        //    else if (dispExtension.Text[0] != '.')
        //    {
        //        MessageBox.Show($"Invalid extension type provided. Extension should start with '.'", "Error");
        //        return;
        //    }
        //    #endregion

        //    #region Get Parameters
        //    string directoryPath = dispDirectory.Text;
        //    Workbook activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Worksheet activeSheet = activeBook.ActiveSheet;
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
        //    #endregion

        //    #region Call method to get files and folders
        //    List<string> files = new List<string>();
        //    getSpecificFiles(directoryPath, dispExtension.Text, ref files, checkNestedFolders.Checked);
        //    #endregion

        //    #region Print results
        //    // Print files array
        //    string[] folder_name = new string[files.Count()];
        //    string[] file_name = new string[files.Count()];
        //    string[] full_path = new string[files.Count()];
        //    int i = 0;
        //    foreach (string file in files)
        //    {
        //        full_path[i] = file;
        //        if (addExtensionCheck.Checked) { file_name[i] = Path.GetFileName(file); }
        //        else { file_name[i] = Path.GetFileNameWithoutExtension(file); }
        //        folder_name[i] = Path.GetFileName(Path.GetDirectoryName(file));
        //        i++;
        //    }
        //    try
        //    {
        //        WriteToExcel(0, 0, true, full_path, folder_name, file_name);
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message == "Nothing found to print")
        //        {
        //            MessageBox.Show("No results found", "Error");
        //        }
        //        else
        //        {
        //            MessageBox.Show($"Error encountered\n\n{ex.Message}", "Error");
        //        }
        //    }
        //    #endregion
        //}

        //private void renameFiles_Click(object sender, EventArgs e)
        //{
        //    #region Check Input Size
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
        //    try { CheckRangeSize(selectedRange, 0, 4); }
        //    catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); return; }
        //    #endregion

        //    #region Get Confirmation
        //    if (DialogResult.OK != MessageBox.Show($"Confirm to rename {selectedRange.Rows.Count} files? This cannot be undone.", "Confirmation"))
        //    {
        //        return;
        //    }
        //    #endregion

        //    #region Read Excel Info
        //    ExcelTable thisTable = new ExcelTable(selectedRange, "Selected Table");
        //    thisTable.AddColumn(1, "sourcePaths");
        //    thisTable.AddColumn(4, "newNames");
        //    thisTable.ReadRangeFromTable();

        //    string[] sourcePaths = thisTable.GetColumnFromName("sourcePaths").ConvertRangeToStringArray();
        //    string[] newNames = thisTable.GetColumnFromName("newNames").ConvertRangeToStringArray();
        //    #endregion

        //    #region Change Names
        //    string[] status = new string[sourcePaths.Length];
        //    int failures = 0;
        //    for (int i = 0; i < sourcePaths.Length; i++)
        //    {
        //        try
        //        {
        //            string sourcePath = sourcePaths[i];
        //            string newName = newNames[i];
        //            if (newName == "")
        //            {
        //                throw new Exception("Error: File Name cannot be empty");
        //            }
        //            string folder = Path.GetDirectoryName(sourcePath);
        //            string newPath = Path.Combine(folder, newName);

        //            status[i] = renameOneFile(sourcePath, newPath);
        //        }
        //        catch (Exception ex)
        //        {
        //            status[i] = "Error: " + ex.Message;
        //        }
        //        if (status[i] != "Completed: File renamed")
        //        {
        //            failures++;
        //        }
        //    }
        //    #endregion

        //    if (failures == 0)
        //    {
        //        MessageBox.Show("Rename operation completed.\n" +
        //             $"{sourcePaths.Length - failures}/{sourcePaths.Length} files renamed", "Completed");
        //    }
        //    else
        //    {
        //        CommonUtilities.WriteToExcel(0, 4, false, status);
        //        MessageBox.Show("Rename operation incomplete.\n" +
        //             $"{sourcePaths.Length-failures}/{sourcePaths.Length} files renamed. Check status.", "Completed");
        //    }
        //}

        //private string renameOneFile(string sourcePath, string newPath)
        //{
        //    #region Check if Path Exist
        //    if (!File.Exists(sourcePath))
        //    {
        //        //MessageBox.Show($"The following file does not exist\n{sourcePath}", "Error");
        //        //throw new Exception($"The following file does not exist\n{sourcePath}", "Error");
        //        return "Error: File does not exist";
        //    }
        //    #endregion

        //    #region Check Extension
        //    if (!Path.HasExtension(newPath))
        //    {
        //        newPath += Path.GetExtension(sourcePath);
        //    }
        //    else if (Path.GetExtension(sourcePath) != Path.GetExtension(newPath))
        //    {
        //        //MessageBox.Show("Inconsistent extension type.\n" +
        //        //    $"Original extension is {Path.GetExtension(sourcePath)} but new extension is {Path.GetExtension(newPath)}.\n" +
        //        //    "Source Path:\n" +
        //        //    $"{sourcePath}");
        //        return "Warning: Inconsistent extension type";
        //    }
        //    #endregion

        //    try
        //    {
        //        File.Move(sourcePath, newPath);
        //        return "Completed: File renamed";
        //    }
        //    catch (Exception ex)
        //    {
        //        return "Error: " + ex.Message;
        //    }
        //}

        //private void insertRenameHeader_Click(object sender, EventArgs e)
        //{
        //    List<string> headers = new List<string>{"File Path", "Folder", "File Name" , "New File Name", "Status"};
        //    InsertHeadersAtSelection(headers, "cols");
        //}
        //#region Recursive Function to get all file paths
        //private void getFilesAndDirectories(string directory, ref List<string> globalFileList, bool checkNest = true)
        //{
        //    // Get all directories and files within the specified directory
        //    string[] subDirectoryList = Directory.GetDirectories(directory);
        //    string[] fileList = Directory.GetFiles(directory);

        //    // Add directories and files to the global lists
        //    globalFileList.AddRange(fileList);

        //    // Recursively call this method for each subdirectory
        //    if (checkNest)
        //    {
        //        foreach (string subDir in subDirectoryList)
        //        {
        //            getFilesAndDirectories(subDir, ref globalFileList);
        //        }
        //    }
        //}

        //private void getFolderDirectories(string directory, ref List<string> globalDirectoryList, bool checkNest = true)
        //{
        //    // Get all directories and files within the specified directory
        //    string[] subDirectoryList = Directory.GetDirectories(directory);

        //    // Recursively call this method for each subdirectory
        //    if (checkNest)
        //    {
        //        foreach (string subDir in subDirectoryList)
        //        {
        //            globalDirectoryList.Add(subDir);
        //            getFolderDirectories(subDir, ref globalDirectoryList);
        //        }
        //    }
        //    else
        //    {
        //        globalDirectoryList.AddRange(subDirectoryList);
        //    }
        //}

        //private void getSpecificFiles(string directory, string extensionType, ref List<string> globalFileList, bool checkNest = true)
        //{
        //    // Get all directories and files within the specified directory
        //    string[] subDirectoryList = Directory.GetDirectories(directory);
        //    string[] fileList = Directory.GetFiles(directory);

        //    // Add directories and files to the global lists
        //    foreach (string file in fileList)
        //    {
        //        Path.GetExtension(file);
        //        if (Path.GetExtension(file) == extensionType)
        //        {
        //            globalFileList.Add(file);
        //        }
        //    }

        //    // Recursively call this method for each subdirectory
        //    if (checkNest)
        //    {
        //        foreach (string subDir in subDirectoryList)
        //        {
        //            getSpecificFiles(subDir, extensionType, ref globalFileList);
        //        }
        //    }
        //}

        //#endregion

        //private static void WriteToExcel(int rowOff, int colOff, bool setCellToText, params Array[] arrays)
        //{
        //    // This code takes any number of arrays (of various types) and outputs them into excel 
        //    // Output order depends on order of the input array
        //    // Output location is the first cell of the current selection, offset by rowOff and colOff

        //    // Find number of rows and columns
        //    int numRow = 0;
        //    int numCol = arrays.Length;
        //    for (int col = 0; col < arrays.Length; col++)
        //    {
        //        if (arrays[col].Length > numRow)
        //        {
        //            numRow = arrays[col].Length; // Finds max number of rows out of all the various arrays
        //        }
        //    }

        //    #region Check if data exist
        //    if (numRow == 0)
        //    {
        //        throw new Exception ("Nothing found to print");
        //    }
        //    #endregion

        //    #region Set Excel Params
        //    // Add section to read input data from Excel
        //    Workbook activeWB = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Worksheet activeWorkSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
        //    Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

        //    // Write to Excel
        //    activeWB.Application.ScreenUpdating = false;

        //    // Write the entire array to the worksheet in one go using Value2
        //    Range startCell = activeWorkSheet.Cells[selectedRange.Row + rowOff, selectedRange.Column + colOff];
        //    Range endCell = startCell.Offset[numRow - 1, numCol - 1];
        //    Range writeRange = activeWorkSheet.Range[startCell, endCell];
        //    #endregion

        //    #region Set cell formatting to text
        //    if (setCellToText)
        //    {
        //        for (int col = 0; col < arrays.Length; col++)
        //        {
        //            if (arrays[col] is string[])
        //            {
        //                Range locStartCell = startCell.Offset[0, col];
        //                Range locEndCell = locStartCell.Offset[numRow-1, 0];
        //                Range formatCell = activeWorkSheet.Range[locStartCell, locEndCell];
        //                formatCell.NumberFormat = "@";
        //            }
        //        }
        //    }
        //    #endregion

        //    // Initiate object
        //    object[,] dataArray = new object[numRow, numCol];
        //    for (int col = 0; col < arrays.Length; col++)
        //    {
        //        for (int row = 0; row < arrays[col].Length; row++)
        //        {
        //            dataArray[row, col] = arrays[col].GetValue(row);
        //        }
        //    }

        //    writeRange.Value2 = dataArray;

        //    activeWB.Application.ScreenUpdating = true;
        //    activeWorkSheet = null;
        //}
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
                for (int index = 0; index < inputDocument.PageCount; index++)
                {
                    PdfPage page = inputDocument.Pages[index];
                    outputDocument.AddPage(page);
                }

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
        private void advancedMerge_Click(object sender, EventArgs e)
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
            int targetColNum = excelTable.Keys.Count;
            if ((endCol - startCol + 1) < targetColNum)
            {
                MessageBox.Show($"Number of columns selected should be {targetColNum}, {endCol - startCol + 1} columns found", "Error");
                return;
            }

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
                MessageBox.Show(ex.Message,"Error");
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
                        InsertHeader(page, headers[relRowNum]);
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
                MessageBox.Show("Final PDF file is empty, no file generated","Error");
                return;
            }
            outputDocument.Save(outputPath);
            #endregion

            MessageBox.Show("Operation completed", "Success");
        }

        private void InsertHeader(PdfPage page, string headerText)
        {
            XGraphics gfx = XGraphics.FromPdfPage(page);
            string fontName = "Arial";
            int fontSize = Convert.ToInt32(dispTitleFontSize.Text);
            XFont fontType = new XFont(fontName, fontSize);

            XRect rect = new XRect(page.Width.Value/6, page.Height.Value/2-50, page.Width.Value/6*4, page.Height.Value-60);
            //gfx.DrawRectangle(XBrushes.AliceBlue, rect);
            string textContents = headerText;
            XTextFormatter tf = new XTextFormatter(gfx);
            tf.LayoutRectangle = rect;
            tf.Alignment = XParagraphAlignment.Center;
            tf.DrawString(textContents, fontType, XBrushes.Black, rect, XStringFormats.TopLeft);
        }

        private void insertRefHeader_Click(object sender, EventArgs e)
        {
            List<string> headers = new List<string>
            {
            "Section Title",
            "Start Pg Num",
            "End Pg Num",
            "Total Pg Num",
            "Insert New Page",
            "File Path",
            "Folder Name",
            "FileName"
            };
            InsertHeadersAtSelection(headers);
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
                InsertHeader(page, titleText[index]);
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
            PdfDocument inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
            PdfDocument outputDocument = new PdfDocument();
            foreach (PdfPage page in inputDocument.Pages)
            {
                outputDocument.AddPage(page);
            }

            int currentPageNum = firstPageNum;
            for (int pageNum = skipPages; pageNum < outputDocument.PageCount; pageNum++)
            {
                #region User Params
                int fontSize = Convert.ToInt32(dispFontSize.Text);
                double xOffset = double.Parse(dispOffsetX.Text);
                double yOffset = double.Parse(dispOffsetY.Text);
                #endregion

                PdfPage page = outputDocument.Pages[pageNum];
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
                worker.ReportProgress(ConvertToProgress(pageNum, outputDocument.PageCount));
                if (worker.CancellationPending)
                {
                    return;
                }
                #endregion

                currentPageNum++;
            }
            #endregion
            outputDocument.Save(outputPath);
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


        private void openPdfInOrder_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Delay
                int sleepDelay = Convert.ToInt32(AttributeTextBoxDic["PdfOpenDelay"].GetDoubleFromTextBox() * 1000);
                #endregion

                #region Get files
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "PDF (*.pdf)|*.pdf";
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
    }
}

