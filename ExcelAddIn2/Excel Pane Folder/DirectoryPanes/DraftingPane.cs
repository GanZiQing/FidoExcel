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
            //{
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

            thisAtt = new AttributeTextBox("totalSheetNum_sheetRenum", dispTotalDwgNum, true); // Need to sort this out
            thisAtt.type = "int";
            AttributeTextBoxDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new AttributeTextBox("fontSize_sheetRenum", dispFontSizeSheetNum, true);
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
            //toolTip1.SetToolTip(setRefTitlePage,
            //    "Set reference title page (provide PDF file).\n" +
            //    "Required for Advance Merge and Create Section Title Page");
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
        double[] thisSheetCoord;
        double[] totalSheetCoord;

        #endregion
        private void editFilesSheetNum_Click(object sender, EventArgs e)
        {
            try
            {
                #region Read input data
                Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
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
                //thisSheetCoord = new double[] 
                //{ 
                //    AttributeTextBoxDic["thisSheetX_sheetRenum"].GetDoubleFromTextBox() * 72 / 25.4,
                //    AttributeTextBoxDic["thisSheetY_sheetRenum"].GetDoubleFromTextBox() * 72 / 25.4
                //};

                //totalSheetCoord = new double[]
                //{
                //    AttributeTextBoxDic["totalSheetX_sheetRenum"].GetDoubleFromTextBox() * 72 / 25.4,
                //    AttributeTextBoxDic["totalSheetY_sheetRenum"].GetDoubleFromTextBox() * 72 / 25.4
                //};



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

                #region Add Sheet number
                if (addSheetNumberCheck.Checked)
                {
                    for (int i = 0; i < finalFilePaths.Length; i++)
                    {
                        AddSheetNumberToOne(finalFilePaths[i], fileNum[i], dispTotalDwgNum.Text);
                    }
                }
                #endregion

                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show($"Error:{ex.Message}"); }
            finally
            {
                fontSize = 0;
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
            AddTextBox(page, sheetNum, fontSize, thisSheetCoord[0], thisSheetCoord[1], "Arial", XBrushes.Red, XBrushes.LightPink);
            AddTextBox(page, totalSheetNum, fontSize, totalSheetCoord[0], totalSheetCoord[1], "Arial", XBrushes.Blue, XBrushes.LightBlue);
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
                    break;
                case 180:
                    gfx.RotateTransform(-180);
                    break;
                case 270:
                    gfx.RotateTransform(-270);
                    break;
                default:
                    break;
                    //case 90:
                    //    //gfx.RotateAtTransform(90, new XPoint(page.Width.Value / 2, page.Height.Value / 2));
                    //    //gfx.RotateTransform(-90);
                    //    //gfx.RotateAtTransform(height-90)
                    //    //gfx.TranslateTransform(0, -page.Height.Value);
                    //    break;
                    //default:
                    //    break;
                    //// Rotate gfx back to original orientation
                    //// Rotation is always done about 
                    //case 90:
                    //    gfx.RotateTransform(-90);
                    //    gfx.TranslateTransform(0, -page.Height.Value);
                    //    //bottomLeftPoint = new XPoint(-xCoord, page.Width.Value - yCoord);
                    //    break;
                    //case 180:
                    //    gfx.RotateTransform(-180);
                    //    bottomLeftPoint = new XPoint(-xCoord, -yCoord);
                    //    break;
                    //case 270:
                    //    gfx.RotateTransform(-270);
                    //    bottomLeftPoint = new XPoint(page.Height.Value - xCoord, -yCoord);
                    //    break;
                    //default:
                    //    bottomLeftPoint = new XPoint(xCoord, yCoord);
                    //    break;
                    //case 90:
                    //    // Translate to the bottom-right corner and rotate -90°

                    //    gfx.RotateTransform(-90);
                    //    gfx.TranslateTransform(page.Width.Value, 0);
                    //    break;
                    //case 180:
                    //    // Translate to the top-right corner and rotate -180°
                    //    gfx.TranslateTransform(page.Width.Value, page.Height.Value);
                    //    gfx.RotateTransform(-180);
                    //    break;
                    //case 270:
                    //    // Translate to the top-left corner and rotate -270°
                    //    gfx.TranslateTransform(0, page.Height).Value;
                    //    gfx.RotateTransform(-270);
                    //    break;
                    //case 0:
                    //default:
                    //    // No transformation needed for 0°
                    //    break;
            }

            #region For Testing only
            for (int x = -2500; x < 2500; x += 100)
            {
                for (int y = -2500; y < 2500; y += 100)
                {
                    bottomLeftPoint = new XPoint(x, y);
                    textContents = $"{x},{y}";
                    XSize textSize = gfx.MeasureString(textContents, fontType);
                    XPoint topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - (textSize.Height));
                    XRect rect = new XRect(topRightPoint, bottomLeftPoint);
                    gfx.DrawRectangle(rectColor, rect);
                    gfx.DrawString(textContents, fontType, fontColor, rect, XStringFormats.BottomRight);
                }
            }
            #endregion
            gfx.Dispose();
            //bottomLeftPoint = new XPoint(xCoord, yCoord);

            //XSize textSize = gfx.MeasureString(textContents, fontType);
            //XPoint topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - (textSize.Height));
            //XRect rect = new XRect(topRightPoint, bottomLeftPoint);
            //gfx.DrawRectangle(rectColor, rect);
            //gfx.DrawString(textContents, fontType, fontColor, rect, XStringFormats.BottomRight);
            //gfx.Dispose();

            //bottomLeftPoint = new XPoint(xTotalSheetNum, yTotalSheetNum);
            //textContents = totalSheetNum;
            //textSize = gfx.MeasureString(textContents, fontType);
            //topRightPoint = new XPoint(bottomLeftPoint.X + textSize.Width, bottomLeftPoint.Y - textSize.Height);
            //rect = new XRect(topRightPoint, bottomLeftPoint);
            //gfx.DrawRectangle(XBrushes.LightPink, rect);
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
    }
}
