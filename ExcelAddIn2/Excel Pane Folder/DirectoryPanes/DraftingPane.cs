using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
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
using System.Windows.Forms.VisualStyles;
using PdfSharp.UniversalAccessibility.Drawing;
using PdfSharp.Fonts;


namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class DraftingPane : UserControl
    {
        #region Initialisers
        Dictionary<string, AttributeTextBox> AttributeTextBoxDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> CustomAttributeDic = new Dictionary<string, CustomAttribute>();

        public DraftingPane()
        {
            InitializeComponent();
            CreateAttributes();
            AddToolTips();
            AddHeaders();
        }     
        
        private void AddHeaders()
        {
            //List<string> headers = new List<string>
            //{"Test"}
            //headers = new List<string> { "Output File Name - leave blank for default", "Sheet Name - leave blank to print all", "File Path" };
            //AddHeaderMenuToButton(printWorkbooks, headers);
        }

        private void CreateAttributes()
        {
            AttributeTextBox thisAtt;
            CustomAttribute thisCustomAtt;

            #region Drafting Sheet Number
            thisAtt = new AttributeTextBox("thisSheetX_sheetRenum", dispThisSheetX, true);
            thisAtt.SetDefaultValue("725");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("thisSheetY_sheetRenum", dispThisSheetY, true);
            thisAtt.SetDefaultValue("30");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("totalSheetX_sheetRenum", dispTotalSheetX, true);
            thisAtt.SetDefaultValue("750");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("totalSheetY_sheetRenum", dispTotalSheetY, true);
            thisAtt.SetDefaultValue("30");
            thisAtt.type = "double";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("totalSheetNum_sheetRenum", dispTotalDwgNum, true);
            thisAtt.type = "int";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region Font Related

            thisAtt = new AttributeTextBox("fontSize_sheetRenum", dispFontSizeSheetNum, true);
            thisAtt.SetDefaultValue("10");
            thisAtt.type = "int";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisCustomAtt = new ComboBoxAttribute("fontName_sheetRenum", dispFontName, "Arial");
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisAtt = new FileTextBox("fontPath_sheetRenum", dispFontPath, setFontFolder);
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region Add sheet checks
            thisCustomAtt = new CheckBoxAttribute("renameFile_sheetRenum", renameFilesCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisCustomAtt = new CheckBoxAttribute("addSheetNum_sheetRenum", addSheetNumberCheck, true);
            CustomAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            #endregion

            #region Test Coordinate
            thisAtt = new AttributeTextBox("inc_sheetRenum", dispIncrement, true);
            thisAtt.SetDefaultValue("100");
            thisAtt.type = "int";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);
            #endregion
        }

        private void AddToolTips()
        {
            toolTip1.SetToolTip(dispFontPath,
                "Font path is set only on the first time \"Edit Files\" is run for each session.\n" +
                "Restart excel if referenced font path is outdated.");
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

        #region Fixed Params
        int fontSize;
        string fontName;
        double[] thisSheetCoord;
        double[] totalSheetCoord;

        #endregion
        string fontPath;
        private void editFilesSheetNum_Click(object sender, EventArgs e)
        {
            try
            {
                #region Check custom font path
                if (!dispFontName.Text.Equals("Custom")) { } // custom font not used, skip check
                else if (dispValidCustomFont.Text.Equals("Custom Font Path: Not set")) { } // custom font not set, skip check
                else
                {
                    string path = dispValidCustomFont.Text;
                    path = path.Substring(18);
                    if (!path.Equals(dispFontPath.Text))
                    {
                        DialogResult res1 = MessageBox.Show($"Custom file path is set and not equals to what is currently provided. \n" +
                            $"Continue with the following font type?\n" +
                            $"{path}", "Warning", MessageBoxButtons.YesNo);
                        if (res1 != DialogResult.Yes) { MessageBox.Show("Restart excel to reset default font"); return; }
                    }
                }
                #endregion

                #region Read input data
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                CheckRangeSize(selectedRange, 0, 3, "Selected Range");
                string[] filePaths = GetContentsAsStringArray(selectedRange.Columns[1], false);
                string[] fileNames = GetContentsAsStringArray(selectedRange.Columns[2], false);
                string[] fileNum = GetContentsAsStringArray(selectedRange.Columns[3], false);

                //int fontSize = Convert.ToInt32(dispFontSizeSheetNum.Text);
                //double xSheetNum = double.Parse(dispThisSheetX.Text) * 72 / 25.4;
                //double ySheetNum = double.Parse(dispThisSheetY.Text) * 72 / 25.4;
                //double xTotalSheetNum = double.Parse(dispTotalSheetX.Text) * 72 / 25.4;
                //double yTotalSheetNum = double.Parse(dispTotalSheetY.Text) * 72 / 25.4;
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

                #region Set Fixed Params
                fontSize = AttributeTextBoxDic["fontSize_sheetRenum"].GetIntFromTextBox();
                fontName = (string)CustomAttributeDic["fontName_sheetRenum"].attValue;
                thisSheetCoord = new double[]
                {
                    AttributeTextBoxDic["thisSheetX_sheetRenum"].GetDoubleFromTextBox(),
                    AttributeTextBoxDic["thisSheetY_sheetRenum"].GetDoubleFromTextBox()
                };
                totalSheetCoord = new double[]
                {
                    AttributeTextBoxDic["totalSheetX_sheetRenum"].GetDoubleFromTextBox(),
                    AttributeTextBoxDic["totalSheetY_sheetRenum"].GetDoubleFromTextBox()
                };
                #endregion

                #region Check Font
                fontPath = dispFontPath.Text;
                GlobalFontSettings.FontResolver = new CustomFontResolver(ref fontPath, ref dispValidCustomFont);
                #endregion

                #region Add Sheet number
                if (addSheetNumberCheck.Checked)
                {
                    for (int i = 0; i < finalFilePaths.Length; i++)
                    {
                        AddSheetNumberToOne(finalFilePaths[i], fileNum[i], dispTotalDwgNum.Text);
                    }
                }
                #endregion
                DialogResult res = MessageBox.Show("Open output folder?", "Completed", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes) { openOutputFolder_Click(sender, e); }
            }
            catch (Exception ex) { MessageBox.Show($"Error:{ex.Message}"); }
            finally
            {
                fontSize = 0;
                fontName = null;
                thisSheetCoord = null;
                totalSheetCoord = null;
            }
        }

        private void AddSheetNumberToOne(string filePath, string sheetNum, string totalSheetNum)
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
            
            #endregion
            
            #region Print Sheet Name
            AddTextBox(page, sheetNum, fontSize, thisSheetCoord[0], thisSheetCoord[1], fontName);
            AddTextBox(page, totalSheetNum, fontSize, totalSheetCoord[0], totalSheetCoord[1], fontName);
            #endregion

            inputDocument.Save(filePath);
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
        
        private void AddTextBox(PdfPage page, string textContents, int fontSize, double xCoord, double yCoord, 
            string fontName = "Arial", XSolidBrush fontColor = null, XSolidBrush rectColor = null)
        {
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XFont fontType = new XFont(fontName, fontSize);
            if (fontColor == null) { fontColor = XBrushes.Black; }
            if (rectColor == null) { rectColor = XBrushes.White; }

            #region Rotation
            int rotation = page.Rotate;
            switch (rotation)
            {
                case 90:
                    gfx.RotateTransform(-90);
                    gfx.TranslateTransform(-page.Height.Value, 0);
                    break;
                case 180:
                    gfx.RotateTransform(-180);
                    gfx.TranslateTransform(-page.Width.Value, -page.Height.Value);
                    break;
                case 270:
                    gfx.RotateTransform(-270);
                    gfx.TranslateTransform(-(page.Width.Value - page.Height.Value), -page.Height.Value);
                    break;
                default:
                    break;
            }
            #endregion
            
            XPoint bottomLeftPoint = new XPoint(xCoord, yCoord);

            XSize textSize = gfx.MeasureString(textContents, fontType);
            XPoint topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - (textSize.Height));
            XRect rect = new XRect(topRightPoint, bottomLeftPoint);
            gfx.DrawRectangle(rectColor, rect);
            gfx.DrawString(textContents, fontType, fontColor, rect, XStringFormats.BottomRight);
            gfx.Dispose();
        }
        #endregion



        private void openOutputFolder_Click(object sender, EventArgs e)
        {
            try
            {
                string workbookPath = Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
                string path = Path.Combine(workbookPath, "Updated Dwg");
                if (!Directory.Exists(path)) { throw new Exception($"Directory does not exist\n{path}"); }
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ex) { MessageBox.Show($"Error:{ex.Message}"); }
        }

        #region Add Coordinates
        private void testAddCoordinate_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get File Path
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "PDF (*.pdf)|*.pdf";
                DialogResult res = openFileDialog.ShowDialog();
                if (res != DialogResult.OK) { return; }
                string filePath = openFileDialog.FileName;
                #endregion

                #region Check if file exist and set pdf file
                if (!File.Exists(filePath)) { throw new Exception($"File does not exist at {filePath}"); }
                PdfDocument inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Modify);
                PdfPage page = inputDocument.Pages[0];
                fontSize = AttributeTextBoxDic["fontSize_sheetRenum"].GetIntFromTextBox();
                #endregion

                #region Add Coordinates
                AddCoordinateMatrix(page, fontSize);
                string folder = Path.GetDirectoryName(filePath);
                string saveFilename = Path.GetFileNameWithoutExtension(filePath) + "_withCoords.pdf";
                string savePath = Path.Combine(folder, saveFilename);
                inputDocument.Save(savePath);
                #endregion

                System.Diagnostics.Process.Start(savePath);
            }
            catch (Exception ex) { MessageBox.Show($"Error:{ex.Message}"); }
            finally
            {
                fontSize = 0;
            }
        }

        private void AddCoordinateMatrix(PdfPage page, int fontSize)
        {
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XFont fontType = new XFont("Arial", fontSize);
            XSolidBrush fontColor = XBrushes.Blue;
            XSolidBrush rectColor = XBrushes.AliceBlue;
            #region Rotation
            int rotation = page.Rotate;
            switch (rotation)
            {
                case 90:
                    gfx.RotateTransform(-90);
                    gfx.TranslateTransform(-page.Height.Value, 0);
                    break;
                case 180:
                    gfx.RotateTransform(-180);
                    gfx.TranslateTransform(-page.Width.Value, -page.Height.Value);
                    break;
                case 270:
                    gfx.RotateTransform(-270);
                    gfx.TranslateTransform(-(page.Width.Value - page.Height.Value), -page.Height.Value);
                    break;
                default:
                    break;
            }
            #endregion
            double width = page.Width.Value;
            double height = page.Height.Value;
            
            int increment = AttributeTextBoxDic["inc_sheetRenum"].GetIntFromTextBox();
            int extent = Convert.ToInt32(Math.Ceiling(Math.Max(width, height) / increment) * increment);
            for (int x = - extent; x < extent; x += increment)
            {
                for (int y = -extent; y < extent; y += increment)
                {
                    string textContents = $"{x},{y}";
                    XPoint bottomLeftPoint = new XPoint(x, y);
                    XSize textSize = gfx.MeasureString(textContents, fontType);
                    XPoint topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - (textSize.Height));
                    XRect rect = new XRect(topRightPoint, bottomLeftPoint);
                    gfx.DrawRectangle(rectColor, rect);
                    gfx.DrawString(textContents, fontType, fontColor, rect, XStringFormats.BottomRight);

                    {
                        XPoint startPoint = new XPoint(x, y);
                        XPoint endPoint = new XPoint(x, y);
                        startPoint.Offset(fontSize / 3, 0);
                        endPoint.Offset(-fontSize / 3, 0);
                        gfx.DrawLine(new XPen(fontColor.Color), startPoint, endPoint);

                        startPoint = new XPoint(x, y);
                        endPoint = new XPoint(x, y);
                        startPoint.Offset(0, fontSize / 3);
                        endPoint.Offset(0, -fontSize / 3);
                        gfx.DrawLine(new XPen(fontColor.Color), startPoint, endPoint);
                    }
                }
            }
            gfx.Dispose();
        }

        #endregion
    }

    public class CustomFontResolver : IFontResolver
    {
        string fontPath;
        System.Windows.Forms.TextBox dispValidCustomFont;
        public CustomFontResolver(ref string fontPath, ref System.Windows.Forms.TextBox dispValidCustomFont) 
        { 
            this.fontPath = fontPath;
            this.dispValidCustomFont = dispValidCustomFont;
        }

        public byte[] GetFont(string fontName)
        {
            if (fontName == "Custom")
            {
                if (!File.Exists(fontPath)) { throw new FileNotFoundException($"Font path '{fontPath}' is invalid."); }
                dispValidCustomFont.Text = $"Custom Font Path: {fontPath}";
                return File.ReadAllBytes(fontPath);
            }
            else { throw new Exception($"Font name {fontName} undefined."); } // This should not trigger
            
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (familyName == "Custom") 
            {
                return new FontResolverInfo("Custom");
            }

            var builtInFont = PlatformFontResolver.ResolveTypeface(familyName, isBold, isItalic);
            return builtInFont;
        }
    }
}
