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

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class ReportPane : UserControl
    {
        #region Initialise
        Dictionary<string, AttributeTextBox> TextBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> CustomAttributeDic = new Dictionary<string, CustomAttribute>();
        public ReportPane()
        {
            InitializeComponent();
            CreateAttributes();
        }

        private void CreateAttributes()
        {
            #region SC Directory
            DirectoryTextBox scFolderPath_report = new DirectoryTextBox("scFolderPath_report", dispSCFolder, setSCFolder);
            scFolderPath_report.AddOpenButton(openSCFolder);
            TextBoxAttributeDic.Add(scFolderPath_report.attName, scFolderPath_report);
            #endregion

            #region ETABS Screenshot Boundary
            var thisAtt = new AttributeTextBox("scX_report", dispScreenshotX, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("30");
            TextBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("scY_report", dispScreenshotY, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("70");
            TextBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("scWidth_report", dispScreenshotWidth, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("800");
            TextBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("scHeight_report", dispScreenshotHeight, true);
            thisAtt.type = "int";
            thisAtt.SetDefaultValue("450");
            TextBoxAttributeDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region ETABS
            RangeTextBox etabsRunRange_report = new RangeTextBox("etabsRunRange_report", dispFloorRange, setFloorRange, "range", true);
            TextBoxAttributeDic.Add("etabsRunRange_report", etabsRunRange_report);

            thisAtt = new AttributeTextBox("startDelay_report", dispStartDelay, true);
            thisAtt.SetDefaultValue("5");
            thisAtt.type = "double";
            TextBoxAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("loadDelay_report", dispLoadDelay, true);
            thisAtt.SetDefaultValue("2");
            thisAtt.type = "double";
            TextBoxAttributeDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region Directory
            //DirectoryTextBox FolderPath = new DirectoryTextBox("FolderPath", dispDirectory, setDirectory);
            //FolderPath.AddOpenButton(dirOpenPath);
            //TextBoxAttributeDic.Add("FolderPath", FolderPath);
            //AttributeTextBox ExtensionType = new AttributeTextBox("ExtensionType", dispExtension, true);
            directoryUserControl1.CreateAttributes(ref TextBoxAttributeDic, ref CustomAttributeDic);
            #endregion

            #region Image Boundary
            AttributeTextBox insertX_report = new AttributeTextBox("insertX_report", dispInsertX, true);
            insertX_report.SetDefaultValue("30");
            TextBoxAttributeDic.Add("insertX_report", insertX_report);

            AttributeTextBox insertY_report = new AttributeTextBox("insertY_report", dispInsertY, true);
            insertY_report.SetDefaultValue("70");
            TextBoxAttributeDic.Add("insertY_report", insertY_report);

            AttributeTextBox widthX_report = new AttributeTextBox("widthX_report", dispWidthX, true);
            widthX_report.SetDefaultValue("780");
            TextBoxAttributeDic.Add("widthX_report", widthX_report);

            AttributeTextBox heightY_report = new AttributeTextBox("heightY_report", dispHeightY, true);
            heightY_report.SetDefaultValue("500");
            TextBoxAttributeDic.Add("heightY_report", heightY_report);
            #endregion

            #region Ppt Import
            FileTextBox pptFilePath = new FileTextBox("pptFilePath", dispPptFile, setPptFile);
            pptFilePath.AddOpenButton(openPpt, ".pptx");
            TextBoxAttributeDic.Add("pptFilePath", pptFilePath);

            RangeTextBox pptImportRange = new RangeTextBox("pptImportRange", dispImportRange, setImportRange, "range", true);
            TextBoxAttributeDic.Add("pptImportRange", pptImportRange);
            #endregion

        }
        #endregion

        #region Ppt Import
        private void importToPpt_Click(object sender, EventArgs e)
        {
            Beaver.InitializeForWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook);
            #region Read Excel
            Range inputRange = ((RangeTextBox)TextBoxAttributeDic["pptImportRange"]).GetRangeFromFullAddress();
            try
            {
                CheckRangeSize(inputRange, 0, 3);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
            
            string[] text1 = GetContentsAsStringArray(inputRange.Columns[1].Cells, false);
            string[] text2 = GetContentsAsStringArray(inputRange.Columns[2].Cells, false);
            string[] imagePaths = GetContentsAsStringArray(inputRange.Columns[3].Cells, false);
            #endregion

            Ppt.Application pptApp = null;
            Ppt.Presentation activePpt = null;

            #region Open ppt
            string pptPath = dispPptFile.Text;
            try
            {
                pptApp = new Ppt.Application();
                pptApp.Visible = MsoTriState.msoTrue;
                activePpt = pptApp.Presentations.Open(pptPath);
            }
            catch (Exception ex)
            {
                if (activePpt != null)
                {
                    activePpt.Close();
                    Marshal.ReleaseComObject(activePpt);
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }

                MessageBox.Show($"Unable to access ppt file" + ex.Message, "Error");
                //throw new Exception($"Unable to access ppt file" + ex.Message);
                return;
            }
            #endregion

            #region Loop Through Excel
            for (int rowNum = 0; rowNum < imagePaths.Length; rowNum++)
            {
                #region Checks and Assignments
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

                #region Add to ppt
                // Add a slide to the presentation
                Ppt.Slide active_slide = activePpt.Slides[1].Duplicate()[1];
                active_slide.MoveTo(activePpt.Slides.Count); // Move to back 

                // Add image 
                float[] imagePosition = GetImageLocation(imagePath);
                active_slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, imagePosition[0], imagePosition[1], imagePosition[2], imagePosition[3]);

                // Change Text Box
                Ppt.Shape text_box = null;
                try
                {
                    text_box = active_slide.Shapes["Description"];
                }
                catch (COMException ex)
                {
                    MessageBox.Show($"Text Box named 'Description' not found, create it in the template and name it using selection pane.\n\n" +
                        $"{ex.Message}" +
                        $"", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (text_box != null)
                {
                    text_box.TextFrame.TextRange.Text = $"{text1[rowNum]} \n{text2[rowNum]}"; // Change the text
                }
                #endregion
            }
            MessageBox.Show("Completed", "Completed");
            Beaver.CheckLog();
            #endregion
        }

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
                TextBoxAttributeDic["insertX_report"].SetValue(Math.Ceiling(rect.Left).ToString());
                TextBoxAttributeDic["insertY_report"].SetValue(Math.Ceiling(rect.Top).ToString());
                TextBoxAttributeDic["widthX_report"].SetValue(Math.Ceiling(rect.Width).ToString());
                TextBoxAttributeDic["heightY_report"].SetValue(Math.Ceiling(rect.Height).ToString());

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
                //MessageBox.Show($"Unable to access ppt file" + ex.Message, "Error");
                throw new Exception($"Unable to access ppt file:\n{pptPath}\n" + ex.Message);
            }
            #endregion
        }

        private void ClosePpt(ref Ppt.Application pptApp, ref Ppt.Presentation activePpt)
        {
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
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
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
                insertX = TextBoxAttributeDic["insertX_report"].GetFloatFromTextBox();
                insertY = TextBoxAttributeDic["insertY_report"].GetFloatFromTextBox();
                inputWidth = TextBoxAttributeDic["widthX_report"].GetFloatFromTextBox();
                inputHeight = TextBoxAttributeDic["heightY_report"].GetFloatFromTextBox();
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
            float finalInsertX = insertX + inputWidth / 2 - finalWidth/2;
            float finalInsertY = insertY + inputHeight / 2 - finalHeight / 2;
            #endregion

            return new float[] { finalInsertX, finalInsertY, finalWidth, finalHeight };
        }


        #endregion

        #endregion

        #region ETABS
        private void saveEtabsImage_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                #region Get Attributes
                int loadDelay;
                int startDelay;
                try
                {
                    double startDelayD = TextBoxAttributeDic["startDelay_report"].GetDoubleFromTextBox();
                    startDelay = Convert.ToInt32(startDelayD * 1000); // Convert from seconds

                    double loadDelayD = TextBoxAttributeDic["loadDelay_report"].GetDoubleFromTextBox();
                    loadDelay = Convert.ToInt32(loadDelayD * 1000); // Convert from seconds
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    return;
                }
                #endregion

                #region Get Excel Data
                Range runRange;
                try
                {
                    runRange = ((RangeTextBox)TextBoxAttributeDic["etabsRunRange_report"]).GetRangeFromFullAddress();
                    CheckRangeSize(runRange, 0, 2, "etabsRunRange_report");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    return;
                }
                string[] toPrint = GetContentsAsStringArray(runRange.Columns[1].Cells, false);
                string[] fileNames = GetContentsAsStringArray(runRange.Columns[2].Cells, false);

                #endregion

                #region Get Screnshot Dimensions and Path
                string folderPath;
                int[] dimensions;
                try
                {
                    folderPath = ((DirectoryTextBox)TextBoxAttributeDic["scFolderPath_report"]).CheckAndGetPath();
                    dimensions = GetScreenshotBoundsFromTextBox();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    return;
                }


                
                #endregion

                #region Check
                progressTracker.UpdateStatus($"Warning Message");
                DialogResult result = MessageBox.Show(
                "WARNING: This will trigger a series of keyboard inputs, even in other softwares.\n" +
                $"Do you want to proceed?\nYou have {startDelay/1000}s to go to the right software.",
                "Warning", MessageBoxButtons.YesNo);

                if (result == DialogResult.No) { return; }
                progressTracker.UpdateStatus($"Pause for {startDelay/ 1000}s");
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

                    worker.ReportProgress(ConvertToProgress(rowNum+1, toPrint.Length));
                    if (worker.CancellationPending)
                    {
                        return;
                    }
                }
                #endregion

                MessageBox.Show("Completed", "Completed");
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
            boundary[0] = TextBoxAttributeDic["scWidth_report"].GetIntFromTextBox();
            boundary[1] = TextBoxAttributeDic["scHeight_report"].GetIntFromTextBox();
            boundary[2] = TextBoxAttributeDic["scX_report"].GetIntFromTextBox();
            boundary[3] = TextBoxAttributeDic["scY_report"].GetIntFromTextBox();

            boundsForm.FormClosed += BoundsForm_FormClosed;
            boundsForm.CloseButton.Text = "Cancel";
            boundsForm.Show();
            boundsForm.SetAllBounds(boundary);
        }

        private void getScreenshotBounds_Click(object sender, EventArgs e)
        {
            if (boundsForm == null)
            {
                MessageBox.Show($"No instance of bounds application found.","Error");
                return;
            }

            int[] boundary = boundsForm.GetAdjustedDimensions();
            TextBoxAttributeDic["scWidth_report"].SetValue(boundary[0].ToString());
            TextBoxAttributeDic["scHeight_report"].SetValue(boundary[1].ToString());
            TextBoxAttributeDic["scX_report"].SetValue(boundary[2].ToString());
            TextBoxAttributeDic["scY_report"].SetValue(boundary[3].ToString());
            
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
                dimensions[0] = TextBoxAttributeDic["scWidth_report"].GetIntFromTextBox();
                dimensions[1] = TextBoxAttributeDic["scHeight_report"].GetIntFromTextBox();
                dimensions[2] = TextBoxAttributeDic["scX_report"].GetIntFromTextBox();
                dimensions[3] = TextBoxAttributeDic["scY_report"].GetIntFromTextBox();
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
                folderPath = ((DirectoryTextBox)TextBoxAttributeDic["scFolderPath_report"]).CheckAndGetPath();
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

