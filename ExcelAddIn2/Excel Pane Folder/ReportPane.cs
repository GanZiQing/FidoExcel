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
using static ExcelAddIn2.CommonUtilities;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Controls.Primitives;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class ReportPane : UserControl
    {
        #region Initialise
        Dictionary<string, AttributeTextBox> textBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> customAttributeDic = new Dictionary<string, CustomAttribute>();
        public ReportPane()
        {
            InitializeComponent();
            CreateAttributes();
            AddToolTips();
            AddHeaders();
        }

        private void CreateAttributes()
        {
            #region SC Directory
            DirectoryTextBox scFolderPath_report = new DirectoryTextBox("scFolderPath_report", dispSCFolder, setSCFolder);
            scFolderPath_report.AddOpenButton(openSCFolder);
            textBoxAttributeDic.Add(scFolderPath_report.attName, scFolderPath_report);

            AttributeTextBox tbAtt = new RangeTextBox("folderNameCell_report", dispFolderNameCell, setFolderNameCell, "cell", false);
            textBoxAttributeDic.Add(tbAtt.attName, tbAtt);

            CustomAttribute customAtt = new CheckBoxAttribute("addToFolder_report", addToFolderCheck, false);
            customAttributeDic.Add(customAtt.attName, customAtt);
            #endregion

            #region ETABS Screenshot Boundary
            var thisAtt = new AttributeTextBox("scX_report", dispScreenshotX, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("30");
            textBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("scY_report", dispScreenshotY, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("70");
            textBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("scWidth_report", dispScreenshotWidth, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("800");
            textBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("scHeight_report", dispScreenshotHeight, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("450");
            textBoxAttributeDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region ETABS
            RangeTextBox etabsRunRange_report = new RangeTextBox("etabsRunRange_report", dispFloorRange, setFloorRange, "range", false);
            textBoxAttributeDic.Add("etabsRunRange_report", etabsRunRange_report);

            thisAtt = new AttributeTextBox("startDelay_report", dispStartDelay, true);
            thisAtt.SetDefaultValue("5");
            thisAtt.type = "double";
            textBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("loadDelay_report", dispLoadDelay, true);
            thisAtt.SetDefaultValue("2");
            thisAtt.type = "double";
            textBoxAttributeDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region Directory
            //DirectoryTextBox FolderPath = new DirectoryTextBox("FolderPath", dispDirectory, setDirectory);
            //FolderPath.AddOpenButton(dirOpenPath);
            //TextBoxAttributeDic.Add("FolderPath", FolderPath);
            //AttributeTextBox ExtensionType = new AttributeTextBox("ExtensionType", dispExtension, true);
            directoryUserControl1.CreateAttributes(ref textBoxAttributeDic, ref customAttributeDic);
            #endregion

            #region Image Boundary
            AttributeTextBox insertX_report = new AttributeTextBox("insertX_report", dispInsertX, true);
            insertX_report.SetDefaultValue("30");
            textBoxAttributeDic.Add("insertX_report", insertX_report);

            AttributeTextBox insertY_report = new AttributeTextBox("insertY_report", dispInsertY, true);
            insertY_report.SetDefaultValue("70");
            textBoxAttributeDic.Add("insertY_report", insertY_report);

            AttributeTextBox widthX_report = new AttributeTextBox("widthX_report", dispWidthX, true);
            widthX_report.SetDefaultValue("780");
            textBoxAttributeDic.Add("widthX_report", widthX_report);

            AttributeTextBox heightY_report = new AttributeTextBox("heightY_report", dispHeightY, true);
            heightY_report.SetDefaultValue("500");
            textBoxAttributeDic.Add("heightY_report", heightY_report);
            #endregion

            #region Ppt Import
            FileTextBox pptFilePath = new FileTextBox("pptFilePath", dispPptFile, setPptFile);
            pptFilePath.AddOpenButton(openPpt, ".pptx");
            textBoxAttributeDic.Add("pptFilePath", pptFilePath);

            tbAtt = new RangeTextBox("pptImportRange", dispImportRange, setImportRange, "range", false);
            textBoxAttributeDic.Add(tbAtt.attName, tbAtt);

            tbAtt = new RangeTextBox("pptImportHeaderRow", dispHeaderRow, setHeaderRow, "row", false);
            textBoxAttributeDic.Add(tbAtt.attName, tbAtt);

            tbAtt = new AttributeTextBox("pptInsertLoc", dispImageLoc, "-1", true);
            tbAtt.type = "int";
            textBoxAttributeDic.Add(tbAtt.attName, tbAtt);

            customAtt = new CheckBoxAttribute("delRefCheck", deleteRefCheck, false);
            customAttributeDic.Add(customAtt.attName, customAtt);
            #endregion

        }
        
        private void AddToolTips()
        {
            ToolTip toolTip1 = new ToolTip();

            #region Ppt File Def
            toolTip1.SetToolTip(setPptFile,
                "Defines reference ppt file, assumes that slide 1 is the template slide"
            );
            #endregion

            #region Image Position
            toolTip1.SetToolTip(insertImageBox,
                "Opens reference ppt and insert box representing the bounding region of the image to be inserted"
            );
            toolTip1.SetToolTip(getBounds,
                "Gets boundary from open reference ppt (from \"Insert Image Box\""
            );
            #endregion

            #region Import to Ppt
            toolTip1.SetToolTip(setImportRange,
                "Takes input in the following format:\n" +
                "  File Path\n" +
                "  Folder Name (unused)\n" +
                "  File Name (unused)\n" +
                "  *ShapeName 1\n" +
                "  *...\n" +
                "* Text Box/Shape Name in ppt (min 1)"
            );

            toolTip1.SetToolTip(setHeaderRow,
                "Defines header row number for import range"
            );

            toolTip1.SetToolTip(deleteRefCheck,
                "If checked, first slide (reference slide) will be deleted"
            );

            toolTip1.SetToolTip(dispImageLoc,
                "Int that defines where the new slide will be inserted.\n" +
                "  -1 will insert slide at the end of the ppt\n" +
                "  0 will insert at the start of the slide\n" +
                "  If location > total number of slides, insert slide at the end of the ppt"
            );
            #endregion
        }

        private void AddHeaders()
        {
            List<string> headers = new List<string> { "File Path", "Folder Name", "File Name", "ShapeName 1", "ShapeName 2" };
            AddHeaderMenuToButton(setImportRange, headers);

            ////Add File Details from Dialogue
            //AddContextStripEvent(importFilePath, "Get From Dialogue Box", (sender, e) => importFilePath_Click(sender, e));
        }
        #endregion

        #region Ppt Import


        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        public void BringExcelToFront()
        {
            var hwnd = Globals.ThisAddIn.Application.Hwnd;
            SetForegroundWindow((IntPtr)hwnd);
        }

        public void BringPptToFront(Ppt.Application pptApp)
        {
            var hwnd = new IntPtr(pptApp.HWND);
            SetForegroundWindow(hwnd);
        }

        private void importToPpt_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                Ppt.Application pptApp = null;
                Ppt.Presentation activePpt = null;
                try
                {
                    progressTracker.UpdateStatus($"Initialising...");
                    Beaver.InitializeForWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);

                    #region Read Excel
                    Range inputRange = ((RangeTextBox)textBoxAttributeDic["pptImportRange"]).GetRangeForCurrentSheet();
                    CheckRangeSize(inputRange, 0, 4, "Import Range", true);
                    string[] imagePaths = GetContentsAsStringArray(inputRange.Columns[1].Cells, false);

                    #region Check and Get Shape Names
                    string[] shapeNames = null;
                    HashSet<string> shapeNamesSet = new HashSet<string>();
                    {
                        Range headerRange = ((RangeTextBox)textBoxAttributeDic["pptImportHeaderRow"]).GetRangeForCurrentSheet();
                        int headerRowNum = headerRange.Row;
                        int inputRangeColNum = inputRange.Column;
                        Range startCell = inputRange.Worksheet.Cells[headerRowNum, inputRangeColNum + 3];
                        Range endRange = inputRange.Worksheet.Cells[headerRowNum, inputRangeColNum + inputRange.Columns.Count - 1];
                        Range shapeNameRange = inputRange.Worksheet.Range[startCell, endRange];
                        //MessageBox.Show($"header range address = {textBoxNameRange.Address}");
                        shapeNames = GetContentsAsStringArray(shapeNameRange, false);
                        foreach (string name in shapeNames)
                        {
                            if (shapeNamesSet.Contains(name)) { continue; }
                            else { shapeNamesSet.Add(name); }
                        }
                    }

                    object[,] shapeContents;
                    {
                        Range shapeContentsRange = inputRange.Offset[0, 3].Resize[inputRange.Rows.Count, inputRange.Columns.Count - 3];
                        shapeContents = GetContentsAsObject2DArray(shapeContentsRange);
                    }
                    #endregion

                    #endregion

                    #region Open ppt
                    string pptPath = dispPptFile.Text;
                    (pptApp, activePpt) = OpenPpt(pptPath);
                    #endregion

                    #region Check Shape Names Exist
                    if (activePpt.Slides.Count < 1) { throw new Exception("Number of slides in reference ppt must be greater than 1"); }
                    Ppt.Slide baseSlide = activePpt.Slides[1];
                    bool[] shapeExist = new bool[shapeNames.Length];

                    foreach (Ppt.Shape shape in baseSlide.Shapes)
                    {
                        if (shapeNamesSet.Contains(shape.Name)) { shapeNamesSet.Remove(shape.Name); }
                    }

                    if (shapeNamesSet.Count > 0)
                    {
                        BringExcelToFront();
                        string shapeNameConcat = "";
                        foreach (string name in shapeNamesSet) { shapeNameConcat += $"\"{name}\", "; }
                        throw new Exception($"Shape(s) named {shapeNameConcat.Substring(0, shapeNameConcat.Length - 2)} not found, create it in the template and name it using selection pane.");
                    }
                    #endregion

                    #region Get Insert Location
                    int insertLoc = ((AttributeTextBox)textBoxAttributeDic["pptInsertLoc"]).GetIntFromTextBox();
                    if (insertLoc > activePpt.Slides.Count || insertLoc < -1) { insertLoc = -1; }
                    #endregion

                    #region Loop Through Excel
                    for (int rowNum = 0; rowNum < imagePaths.Length; rowNum++)
                    {
                        progressTracker.UpdateStatus($"Adding images {rowNum + 1}/{imagePaths.Length}");
                        #region Checks Image Path
                        string imagePath = imagePaths[rowNum];
                        if (imagePaths[rowNum] == "")
                        {
                            continue;
                        }

                        // Check if image file exist
                        if (!File.Exists(imagePaths[rowNum]))
                        {
                            Beaver.LogError($"File path not found at row {rowNum} of input range. Image not added.\n{imagePath}");
                            continue;
                        }
                        #endregion

                        #region Add new slide
                        Ppt.Slide activeSlide = activePpt.Slides[1].Duplicate()[1];
                        if (insertLoc == -1)
                        {
                            activeSlide.MoveTo(activePpt.Slides.Count); // Move to back
                        }
                        else
                        {
                            activeSlide.MoveTo(insertLoc + 1); // Move to back
                            insertLoc++;
                        }
                        #endregion

                        #region Modify Slide
                        // Add image 
                        float[] imagePosition = GetImageLocation(imagePath);
                        activeSlide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, imagePosition[0], imagePosition[1], imagePosition[2], imagePosition[3]);

                        // Change Shape 
                        HashSet<string> modifiedShapeNames = new HashSet<string>();
                        for (int colIndex = 0; colIndex < shapeNames.Length; colIndex++)
                        {
                            string shapeName = shapeNames[colIndex];
                            Ppt.Shape shape = null;
                            try
                            {
                                shape = activeSlide.Shapes[shapeName];
                                if (!modifiedShapeNames.Contains(shapeName))
                                {
                                    // Change the text
                                    shape.TextFrame.TextRange.Text = shapeContents[rowNum, colIndex].ToString();
                                }
                                else
                                {
                                    // Concat text
                                    shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text + "\n" + shapeContents[rowNum, colIndex].ToString();
                                }
                                modifiedShapeNames.Add(shapeName);
                            }
                            catch (Exception ex)
                            {
                                Beaver.LogError($"Unable to change text in shape named {shapeName}\n ErrorMsg: {ex.Message}");
                            }
                        }
                        #endregion

                        worker.ReportProgress(ConvertToProgress(rowNum + 1, imagePaths.Length));
                    }
                    #endregion

                    if (deleteRefCheck.Checked) { baseSlide.Delete(); }
                    BringExcelToFront();
                    Beaver.CheckLog();
                    progressTracker.UpdateStatus($"Completed Message Box Shown");
                    progressTracker.ShowMessageBox("Images added to ppt, please check and save ppt as required.", "Completed");
                    BringPptToFront(pptApp);
                }
                catch (Exception ex)
                {
                    progressTracker.UpdateStatus($"Warning Message Box Shown");
                    BringExcelToFront();
                    progressTracker.ShowMessageBox(ex.Message, "Error");
                }
                finally
                {
                    pptApp = null;
                    activePpt = null;
                }
            });
        }

        //private void importToPpt_Click(object sender, EventArgs e)
        //{
        //    Beaver.InitializeForWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);

        //    #region Read Excel
        //    Range inputRange = ((RangeTextBox)TextBoxAttributeDic["pptImportRange"]).GetRangeFromFullAddress();
        //    try
        //    {
        //        //CheckRangeSize(inputRange, 0, 4, "Import Range", true);
        //        CheckRangeSize(inputRange, 0, 3);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error");
        //        return;
        //    }

        //    string[] text1 = GetContentsAsStringArray(inputRange.Columns[1].Cells, false);
        //    string[] text2 = GetContentsAsStringArray(inputRange.Columns[2].Cells, false);
        //    string[] imagePaths = GetContentsAsStringArray(inputRange.Columns[3].Cells, false);
        //    #endregion

        //    Ppt.Application pptApp = null;
        //    Ppt.Presentation activePpt = null;

        //    #region Open ppt
        //    string pptPath = dispPptFile.Text;
        //    try
        //    {
        //        pptApp = new Ppt.Application();
        //        pptApp.Visible = MsoTriState.msoTrue;
        //        activePpt = pptApp.Presentations.Open(pptPath);
        //    }
        //    catch (Exception ex)
        //    {
        //        if (activePpt != null)
        //        {
        //            activePpt.Close();
        //            Marshal.ReleaseComObject(activePpt);
        //        }
        //        if (pptApp != null)
        //        {
        //            pptApp.Quit();
        //            Marshal.ReleaseComObject(pptApp);
        //        }

        //        MessageBox.Show($"Unable to access ppt file" + ex.Message, "Error");
        //        //throw new Exception($"Unable to access ppt file" + ex.Message);
        //        return;
        //    }
        //    #endregion

        //    #region Loop Through Excel
        //    for (int rowNum = 0; rowNum < imagePaths.Length; rowNum++)
        //    {
        //        #region Checks and Assignments
        //        string imagePath = imagePaths[rowNum];
        //        if (imagePaths[rowNum] == "")
        //        {
        //            continue;
        //        }

        //        // Check if image file exist
        //        if (!File.Exists(imagePaths[rowNum]))
        //        {
        //            Beaver.LogError($"File path not found at row {rowNum} of input range. Image not added.\n{imagePath}");
        //            continue;
        //        }
        //        #endregion

        //        #region Add to ppt
        //        // Add a slide to the presentation
        //        Ppt.Slide active_slide = activePpt.Slides[1].Duplicate()[1];
        //        active_slide.MoveTo(activePpt.Slides.Count); // Move to back 

        //        // Add image 
        //        float[] imagePosition = GetImageLocation(imagePath);
        //        active_slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, imagePosition[0], imagePosition[1], imagePosition[2], imagePosition[3]);

        //        // Change Text Box
        //        Ppt.Shape text_box = null;
        //        try
        //        {
        //            text_box = active_slide.Shapes["Description"];
        //        }
        //        catch (COMException ex)
        //        {
        //            MessageBox.Show($"Text Box named 'Description' not found, create it in the template and name it using selection pane.\n\n" +
        //                $"{ex.Message}" +
        //                $"", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            return;
        //        }
        //        if (text_box != null)
        //        {
        //            text_box.TextFrame.TextRange.Text = $"{text1[rowNum]} \n{text2[rowNum]}"; // Change the text
        //        }
        //        #endregion
        //    }
        //    MessageBox.Show("Completed", "Completed");
        //    Beaver.CheckLog();
        //    #endregion
        //}
        #region Image Boundary
        Ppt.Application pptApp_image = null;
        Ppt.Presentation activePpt_image = null;
        private void insertImageBox_Click(object sender, EventArgs e)
        {
            #region Open ppt
            string pptPath = dispPptFile.Text;
            try
            {
                (pptApp_image, activePpt_image) = OpenPpt(pptPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "Error");
                return;
            }
            #endregion

            #region Get Dimensions
            (float insertX, float insertY, float inputWidth, float inputHeight) = GetInputDimensions();
            #endregion

            #region Insert Rectangle
            Ppt.Slide slide = activePpt_image.Slides[1];
            Ppt.Shape rectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, insertX, insertY, inputWidth, inputHeight);
            rectangle.Name = "Reference Rectangle";
            #endregion
        }

        private void getBounds_Click(object sender, EventArgs e)
        {
            // Check Ppt Existance 
            if (pptApp_image == null || activePpt_image == null)
            {
                MessageBox.Show("No instance of power point found.", "Error");
                return;
            }

            // Get Shape
            //Ppt.Shape rect = activePpt_image.Slides[1].Shapes["Reference Rectangle"];
            try
            {
                Ppt.Shape rect = activePpt_image.Slides[1].Shapes["Reference Rectangle"];
                // Write to properties
                textBoxAttributeDic["insertX_report"].SetValue(Math.Ceiling(rect.Left).ToString());
                textBoxAttributeDic["insertY_report"].SetValue(Math.Ceiling(rect.Top).ToString());
                textBoxAttributeDic["widthX_report"].SetValue(Math.Ceiling(rect.Width).ToString());
                textBoxAttributeDic["heightY_report"].SetValue(Math.Ceiling(rect.Height).ToString());

                MessageBox.Show($"Values set.", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to get boundary\n{ex.Message}", "Error");
            }

            ClosePpt(ref pptApp_image, ref activePpt_image);
        }
        #endregion

        #region Helper Functions
        private (Ppt.Application, Ppt.Presentation) OpenPpt(string pptPath)
        {
            Ppt.Application pptApp = null;
            Ppt.Presentation activePpt = null;

            #region Open ppt
            //if (!File.Exists(pptPath)) { throw new Exception($"Unable to find file located at: {pptPath}"); }
            try
            {
                pptApp = new Ppt.Application();
                pptApp.Visible = MsoTriState.msoTrue;
                activePpt = pptApp.Presentations.Open(pptPath);
                return (pptApp, activePpt);
            }
            catch (Exception ex)
            {
                ClosePpt(ref pptApp, ref activePpt);
                throw new Exception($"Unable to access ppt file:\n{pptPath}\n" + ex.Message);
            }
            #endregion
        }

        private void ClosePpt(ref Ppt.Application pptApp, ref Ppt.Presentation activePpt)
        {
            // This does not actually work to close ppt applications unfortunately :(
            // Not sure why
            try
            {
                if (activePpt != null)
                {
                    activePpt.Close();
                    Marshal.ReleaseComObject(activePpt);
                }
            }
            finally
            {
                activePpt = null;
            }
            try
            {
                if (pptApp != null)
                {
                    int pptCount = pptApp.Presentations.Count;
                    if (pptCount == 0)
                    {
                        pptApp.Quit();
                        Marshal.ReleaseComObject(pptApp);
                    }
                }
            }
            finally
            {
                pptApp = null;
            }
        }

        private int[] FindImageLocation_1(string imagePath, Ppt.Slide activeSlide, double offXP = 0, double offYP = 0, double clearanceX = 0.95, double clearanceY = 0.95)
        {
            // Not in use anymore
            // offXP and offYP represents the offset of the image in terms of % of the entire slide size
            // clearanceX/Y = % of the remaining width & length you want to fill the image by
            // Get ppt size
            double slideWidth = activeSlide.Master.Width; // Width of the slide [X]
            double slideHeight = activeSlide.Master.Height; // Height of the slide [Y]
            double width = 0;
            double height = 0;
            using (Image img = Image.FromFile(imagePath))
            {
                // Get the width and height of the image
                width = img.Width;
                height = img.Height;

                // Set the page dimensions
                double ppt_width = (double)(slideWidth - offXP * slideWidth) * clearanceX; // Offset both sides
                double ppt_height = (double)(slideHeight - offYP * slideHeight) * clearanceY;

                // Find Scale
                double scale = Math.Min(ppt_height / height, ppt_width / width);
                double a = ppt_height / height;
                double b = ppt_width / width;
                double final_height = (height * scale);
                double final_width = (width * scale);

                // Find offset
                int final_offX = (int)(offXP * slideWidth + (slideWidth - final_width) / 2);
                int final_offY = (int)(offYP * slideHeight);
                int[] output = new int[] { final_offX, final_offY, (int)final_width, (int)final_height };
                return output;
            }
        }

        private (float, float, float, float) GetInputDimensions()
        {
            //(float insertX, float insertY, float inputWidth, float inputHeight) = GetInputDimensions();
            #region Get Input Dimensions
            float insertX;
            float insertY;
            float inputWidth;
            float inputHeight;

            try
            {
                insertX = textBoxAttributeDic["insertX_report"].GetFloatFromTextBox();
                insertY = textBoxAttributeDic["insertY_report"].GetFloatFromTextBox();
                inputWidth = textBoxAttributeDic["widthX_report"].GetFloatFromTextBox();
                inputHeight = textBoxAttributeDic["heightY_report"].GetFloatFromTextBox();
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable determine to get input image size, {ex.Message}");
            }
            #endregion
            return (insertX, insertY, inputWidth, inputHeight);
        }

        private float[] GetImageLocation(string imagePath)
        {
            (float insertX, float insertY, float inputWidth, float inputHeight) = GetInputDimensions();

            #region Get Image Dimensions
            float imageWidth;
            float imageHeight;
            using (Image img = Image.FromFile(imagePath))
            {
                imageWidth = img.Width;
                imageHeight = img.Height;
            }
            #endregion

            #region Scale
            float scale = Math.Min(inputWidth / imageWidth, inputHeight / imageHeight);
            float finalHeight = (imageHeight * scale);
            float finalWidth = (imageWidth * scale);
            #endregion

            #region Center image
            float finalInsertX = insertX + inputWidth / 2 - finalWidth / 2;
            float finalInsertY = insertY + inputHeight / 2 - finalHeight / 2;
            #endregion

            return new float[] { finalInsertX, finalInsertY, finalWidth, finalHeight };
        }


        #endregion

        #endregion

        #region ETABS Screenshots
        private void saveEtabsImage_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                try
                {
                    #region Get Attributes
                    int loadDelay;
                    int startDelay;
                    try
                    {
                        double startDelayD = textBoxAttributeDic["startDelay_report"].GetDoubleFromTextBox();
                        startDelay = Convert.ToInt32(startDelayD * 1000); // Convert from seconds

                        double loadDelayD = textBoxAttributeDic["loadDelay_report"].GetDoubleFromTextBox();
                        loadDelay = Convert.ToInt32(loadDelayD * 1000); // Convert from seconds
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error getting Attributes: {ex.Message}");
                        //MessageBox.Show(ex.Message, "Error");
                        //return;
                    }
                    #endregion

                    #region Get Excel Data
                    Range runRange;
                    try
                    {
                        runRange = ((RangeTextBox)textBoxAttributeDic["etabsRunRange_report"]).GetRangeForCurrentSheet();
                        CheckRangeSize(runRange, 0, 2, "etabsRunRange_report");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error getting excel data: {ex.Message}");
                        //MessageBox.Show(ex.Message, "Error");
                        //return;
                    }
                    string[] toPrint = GetContentsAsStringArray(runRange.Columns[1].Cells, false);
                    string[] fileNames = GetContentsAsStringArray(runRange.Columns[2].Cells, false);

                    #endregion

                    #region Get Screnshot Dimensions and Path
                    string folderPath;
                    int[] dimensions;
                    try
                    {
                        folderPath = ((DirectoryTextBox)textBoxAttributeDic["scFolderPath_report"]).CheckAndGetPath();
                        if (addToFolderCheck.Checked) { 
                            Range folderNameCell = ((RangeTextBox)textBoxAttributeDic["folderNameCell_report"]).GetRangeForCurrentSheet();
                            string folderName = GetContentsAsString(folderNameCell);
                            folderPath = Path.Combine(folderPath, folderName);
                            if (!Directory.Exists(folderPath)) { Directory.CreateDirectory(folderPath); }
                        }
                        
                        dimensions = GetScreenshotBoundsFromTextBox();

                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error taking screenshot: {ex.Message}");
                        //MessageBox.Show(ex.Message, "Error");
                        //return;
                    }
                    #endregion

                    #region Check
                    progressTracker.UpdateStatus($"Warning Message");
                    
                    DialogResult result = progressTracker.ShowMessageBox(
                    "WARNING: This will trigger a series of keyboard inputs, even in other softwares.\n" +
                    $"Do you want to proceed?\nYou have {startDelay / 1000}s to go to the right software.",
                    "Warning", MessageBoxButtons.YesNo);

                    if (result == DialogResult.No) { return; }
                    progressTracker.UpdateStatus($"Pause for {startDelay / 1000}s");
                    System.Threading.Thread.Sleep(startDelay); // Wait for 5 seconds
                    #endregion

                    #region Loop
                    for (int rowNum = 0; rowNum < toPrint.Length; rowNum++)
                    {
                        progressTracker.UpdateStatus($"Processing {fileNames[rowNum]}");
                        if (toPrint[rowNum] == "1")
                        {
                            string fileName = fileNames[rowNum];
                            string filePath = Path.Combine(folderPath, fileName + ".png");
                            CaptureScreenOnce(filePath, dimensions);
                        }
                        SendKeys.SendWait("{PGUP}"); // Page Up
                        System.Threading.Thread.Sleep(loadDelay); // Wait

                        worker.ReportProgress(ConvertToProgress(rowNum + 1, toPrint.Length));
                        if (worker.CancellationPending)
                        {
                            return;
                        }
                    }
                    #endregion

                    BringExcelToFront();
                    progressTracker.ShowMessageBox("Completed", "Completed");
                }
                catch (Exception ex) { BringExcelToFront(); progressTracker.ShowMessageBox(ex.Message, "Error"); }
            });
        }

        private void saveEtabsImage_Click_Og(object sender, EventArgs e)
        {
            //ProgressHelper.RunWithProgress((worker, progressTracker) =>
            //{
            //    #region Get Attributes
            //    string keyboardShortcut;
            //    int printDelay;
            //    int noPrintDelay;
            //    int startDelay;
            //    try
            //    {
            //        keyboardShortcut = TextBoxAttributeDic["keyboardShortcut_report"].textBox.Text;

            //        double startDelayD = TextBoxAttributeDic["startDelay_report"].GetDoubleFromTextBox();
            //        startDelay = Convert.ToInt32(startDelayD * 1000); // Convert from seconds

            //        double printDelayD = TextBoxAttributeDic["printDelay_report"].GetDoubleFromTextBox();
            //        printDelay = Convert.ToInt32(printDelayD * 1000); // Convert from seconds

            //        double noPrintDelayD = TextBoxAttributeDic["noPrintDelay_report"].GetDoubleFromTextBox();
            //        noPrintDelay = Convert.ToInt32(noPrintDelayD * 1000); // Convert from seconds
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "Error");
            //        return;
            //    }
            //    #endregion

            //    #region Get Excel Data
            //    Range runRange;
            //    try
            //    {
            //        runRange = ((RangeTextBox)TextBoxAttributeDic["etabsRunRange_report"]).GetRangeFromFullAddress();
            //        CheckRangeSize(runRange, 0, 2, "etabsRunRange_report");
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "Error");
            //        return;
            //    }
            //    string[] toPrint = GetContentsAsStringArray(runRange.Columns[1].Cells, false);
            //    string[] fileNames = GetContentsAsStringArray(runRange.Columns[2].Cells, false);

            //    #endregion

            //    #region Check
            //    progressTracker.UpdateStatus($"Warning Message");
            //    DialogResult result = MessageBox.Show(
            //    "WARNING: This will trigger a series of keyboard inputs, even in other softwares.\n" +
            //    $"Do you want to proceed?\nYou have {startDelay / 1000}s to go to the right software.",
            //    "Warning", MessageBoxButtons.YesNo);

            //    if (result == DialogResult.No) { return; }
            //    progressTracker.UpdateStatus($"Pause for {startDelay / 1000}s");
            //    System.Threading.Thread.Sleep(startDelay); // Wait for 5 seconds
            //    #endregion

            //    #region Loop 
            //    for (int rowNum = 0; rowNum < toPrint.Length; rowNum++)
            //    {
            //        progressTracker.UpdateStatus($"Processing {fileNames[rowNum]}");
            //        if (toPrint[rowNum] == "1")
            //        {
            //            SendKeys.SendWait(keyboardShortcut);
            //            string textOutput = fileNames[rowNum];
            //            SendKeys.SendWait(textOutput);
            //            SendKeys.SendWait("{ENTER}");
            //            System.Threading.Thread.Sleep(printDelay); // Wait
            //        }
            //        else
            //        {
            //            System.Threading.Thread.Sleep(noPrintDelay); // Wait
            //        }
            //        SendKeys.SendWait("{PGUP}"); // Page Up

            //        worker.ReportProgress(ConvertToProgress(rowNum + 1, toPrint.Length));
            //        if (worker.CancellationPending)
            //        {
            //            return;
            //        }
            //    }
            //    #endregion

            //    MessageBox.Show("Completed", "Completed");
            //});
        }
        #endregion

        #region Screenshot Boundary
        private void launchScreenshotApp_Click(object sender, EventArgs e)
        {
            ScreenshotApp.ScreenshotForm standaloneApp = new ScreenshotApp.ScreenshotForm();
            standaloneApp.Show();
        }

        ScreenshotApp.ScreenshotForm boundsForm;
        private void setScreenshotBounds_Click(object sender, EventArgs e)
        {
            boundsForm = new ScreenshotApp.ScreenshotForm(true);
            //boundsForm.TopMost = false;

            int[] boundary = new int[4];
            boundary[0] = textBoxAttributeDic["scWidth_report"].GetIntFromTextBox();
            boundary[1] = textBoxAttributeDic["scHeight_report"].GetIntFromTextBox();
            boundary[2] = textBoxAttributeDic["scX_report"].GetIntFromTextBox();
            boundary[3] = textBoxAttributeDic["scY_report"].GetIntFromTextBox();

            boundsForm.FormClosed += BoundsForm_FormClosed;
            boundsForm.CloseButton.Text = "Cancel";
            boundsForm.Show();
            boundsForm.SetAllBounds(boundary);
        }

        private void getScreenshotBounds_Click(object sender, EventArgs e)
        {
            if (boundsForm == null)
            {
                MessageBox.Show($"No instance of bounds application found.", "Error");
                return;
            }

            int[] boundary = boundsForm.GetAdjustedDimensions();
            textBoxAttributeDic["scWidth_report"].SetValue(boundary[0].ToString());
            textBoxAttributeDic["scHeight_report"].SetValue(boundary[1].ToString());
            textBoxAttributeDic["scX_report"].SetValue(boundary[2].ToString());
            textBoxAttributeDic["scY_report"].SetValue(boundary[3].ToString());

            boundsForm.Close();
            MessageBox.Show("Boundary updated.", "Completed");
        }

        private void BoundsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            boundsForm.Dispose();
            boundsForm = null;
        }

        private int[] GetScreenshotBoundsFromTextBox()
        {
            int[] dimensions = new int[4];
            try
            {
                dimensions[0] = textBoxAttributeDic["scWidth_report"].GetIntFromTextBox();
                dimensions[1] = textBoxAttributeDic["scHeight_report"].GetIntFromTextBox();
                dimensions[2] = textBoxAttributeDic["scX_report"].GetIntFromTextBox();
                dimensions[3] = textBoxAttributeDic["scY_report"].GetIntFromTextBox();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting screenshot dimesions {ex.Message}");

            }
            return dimensions;
        }
        #endregion

        #region Screenshot
        private void testScreenshot_Click(object sender, EventArgs e)
        {
            #region Get File Path 
            string folderPath;
            try
            {
                folderPath = ((DirectoryTextBox)textBoxAttributeDic["scFolderPath_report"]).CheckAndGetPath();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }

            string fileName = "Test Screenshot.png";
            string filePath = Path.Combine(folderPath, fileName);
            #endregion

            #region Get Dimensions
            int[] dimensions;
            try
            {
                dimensions = GetScreenshotBoundsFromTextBox();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
            #endregion

            CaptureScreenOnce(filePath, dimensions);

            #region Open Screenshot
            try
            {
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            #endregion
        }

        private void CaptureScreenOnce(string filePath, int[] dimensions)
        {
            Bitmap captureBitmap = new Bitmap(dimensions[0], dimensions[1], PixelFormat.Format32bppArgb);
            Graphics captureGraphics = Graphics.FromImage(captureBitmap);
            captureGraphics.CopyFromScreen(dimensions[2], dimensions[3], 0, 0, new Size(dimensions[0], dimensions[1]));
            //captureGraphics.CopyFromScreen(0, 0, 0, 0, new Size(dimensions[0], dimensions[1]));

            // Save image
            captureBitmap.Save(filePath, ImageFormat.Png);
        }
        #endregion

    }
}

