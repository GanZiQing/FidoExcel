using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using System.Diagnostics;
using static ExcelAddIn2.CommonUtilities;
using System.Text.RegularExpressions;
using System.Diagnostics.Eventing.Reader;

namespace ExcelAddIn2.Excel_Pane_Folder.HDB_Design
{
    public partial class WallCheckPane : UserControl
    {
        #region Init
        //Dictionary<string, AttributeTextBox> TextBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        //Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();
        Dictionary<string, object> attributeDic = new Dictionary<string, object>();
        public WallCheckPane()
        {
            InitializeComponent();
            CreateAttributes();
            AddToolTips();
        }

        private void CreateAttributes()
        {
            #region Base
            AttributeTextBox attTB = new FileTextBox("etabsOutputFile_WC", dispEtabsOutputFile, setEtabsOutputFile);
            attributeDic.Add(attTB.attName, attTB);
            
            attTB = new FileTextBox("bimOutputFile_WC", dispBimOutputFile, setBimOutputFile);
            attributeDic.Add(attTB.attName, attTB);

            attTB = new SheetTextBox("wallCheckSheet_WC", dispWallCheckSheet, setWallCheckSheet);
            attributeDic.Add(attTB.attName, attTB);

            attTB = new SheetTextBox("colCheckSheet_WC", dispColCheckSheet, setColCheckSheet);
            attributeDic.Add(attTB.attName, attTB);
            #endregion

            #region Optional Parameters
            attTB = new AttributeTextBox("etabsSheetName_WC", dispEtabsShtNm, "Pier Dgn Sum - Eurocode 2-2004");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("bimWallSheetName_WC", dispBimWallShtNm, "RC WALL");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("bimHsSheetName_WC", dispBimHsShtNm, "HS WALL");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("bimColSheetName_WC", dispBimColShtNm, "RC COLUMN");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("checkShtNm_WC", dispCheckShtNm, "Check Sheet");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("storeyShtNm_WC", dispStoreyMapShtNm, "Storey Mapping");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("labelShtNm_WC", dispLabelMapShtNm, "Label Mapping");
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("dispColOuputColNum_WC", dispColOuputColNum, "14");
            attTB.type = "int";
            attributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("dispWallOuputColNum_WC", dispWallOuputColNum, "26");
            attTB.type = "int";
            attributeDic.Add(attTB.attName, attTB);

            var att = new CheckBoxAttribute("copyFailedToClip_WC", copyFailedCheck, true);
            attributeDic.Add(att.attName, att);
            #endregion

        }
        private void AddToolTips()
        {
            //toolTip1.SetToolTip(overwriteRebarCheck,
            //    "If unchecked, initial check will be done based on current values in the output range\n" +
            //    "If checked, initial check will be done based on values matched from Rebar Table");
        }
        public TabControl.TabPageCollection GetPageTaskPane
        {
            get { return tabControl1.TabPages; }
        }
        #endregion

        #region Check Walls
        private void checkWalls_Click(object sender, EventArgs e)
        {
            try
            {
                Stopwatch totalStopwatch = Stopwatch.StartNew();
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                Worksheet worksheet = ((SheetTextBox)attributeDic["wallCheckSheet_WC"]).getSheet();
                worksheet.Activate();

                ReadMapping();
                ReadETABSInput(worksheet);
                ReadBIMInput();
                MatchDesignLabels(totalStopwatch);
                //totalStopwatch.Stop();
                MessageBox.Show($"Total Execution Time: {totalStopwatch.ElapsedMilliseconds} ms", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                storeyMap = null;
                labelMap = null;
                etabsRange = null;
                wallToGroupingDic = null;
            }
        }

        #endregion

        #region Check Cols
        private void checkCols_Click(object sender, EventArgs e)
        {
            try
            {
                Stopwatch totalStopwatch = Stopwatch.StartNew();
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                Worksheet worksheet = ((SheetTextBox)attributeDic["colCheckSheet_WC"]).getSheet();
                worksheet.Activate();

                ReadMapping();
                ReadETABSInput(worksheet);
                ReadBIMInput();
                MatchColDesignLabels(totalStopwatch);
                //totalStopwatch.Stop();
                MessageBox.Show($"Total Execution Time: {totalStopwatch.ElapsedMilliseconds} ms", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                storeyMap = null;
                labelMap = null;
                etabsRange = null;
                wallToGroupingDic = null;
            }
        }
        #endregion

        #region Copy from ETABS
        private void copyFromETABS_Click(object sender, EventArgs e)
        {
            Workbook etabsWorkbook = null;
            bool workbookOpened = false;
            try
            {
                #region Clear current sheet
                Range destinationRange = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range["A4"];
                #endregion

                #region Get Copy Range
                (etabsWorkbook,workbookOpened) = ((FileTextBox)attributeDic["etabsOutputFile_WC"]).OpenAndGetWorkbook(Globals.ThisAddIn.Application);
                string worksheetNm = ((AttributeTextBox)attributeDic["etabsSheetName_WC"]).textBox.Text;
                etabsRange = new ExcelTableRange("etabs", etabsWorkbook, worksheetNm);
                etabsRange.GetUsedRangeFromEnd(2, 1, 2);
                Range copyRange = etabsRange.activeRange;                
                #endregion

                #region Clear
                string copyRangeAddress = copyRange.Address;
                Range startCell = destinationRange.Worksheet.Range["A4"];
                Range endCell = destinationRange.Worksheet.Cells[1048576, copyRange.Column + copyRange.Columns.Count - 1];
                Range clearRange = destinationRange.Worksheet.Range[startCell, endCell];
                DialogResult res = MessageBox.Show($"Clear range: {clearRange.Worksheet.Name}!{clearRange.Address}?","Confirmation",MessageBoxButtons.YesNo);
                if (res != DialogResult.Yes) { throw new Exception("Terminated by user"); }
                clearRange.Clear();
                #endregion

                copyRange.Copy(destinationRange);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                if (etabsWorkbook != null && workbookOpened) { etabsWorkbook.Close(); etabsWorkbook = null; }
            }

        }
        #endregion

        #region Read Reference Files
        EtabsToDesignMap storeyMap;
        EtabsToDesignMap labelMap;
        private void ReadMapping()
        {
            try
            {
                Workbook currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                string mapShtNm = ((AttributeTextBox)attributeDic["storeyShtNm_WC"]).textBox.Text;
                ExcelTableRange storeyMapRange = new ExcelTableRange("storyMap", currentWorkbook, mapShtNm);
                storeyMap = new EtabsToDesignMap("storey", storeyMapRange.GetUsedRangeFromEnd(1, 1, 2), new bool[] { true, true, true });

                string labelShtNm = ((AttributeTextBox)attributeDic["labelShtNm_WC"]).textBox.Text;
                ExcelTableRange labelMapRange = new ExcelTableRange("labelMap", currentWorkbook, labelShtNm);
                labelMap = new EtabsToDesignMap("label", labelMapRange.GetUsedRangeFromEnd(1, 1, 2), new bool[] { true, false, false });
            }
            catch (Exception ex) { throw new Exception("Error reading Mapping Input table\n" + ex.Message); }
        }

        ExcelTableRange etabsRange;
        private void ReadETABSInput(Worksheet worksheet)
        {
            try
            {
                //Range destinationRange = worksheet.Range["A4"];
                etabsRange = new ExcelTableRange("etabs", worksheet);
                etabsRange.GetUsedRangeFromStart(2, 1, 2);
            }
            catch (Exception ex) { throw new Exception("Error reading ETABS Input table\n" + ex.Message); }
        }

        Dictionary<string, DesignGroupBIM> designGroupsDic;
        Dictionary<string, string> wallToGroupingDic;
        private void ReadBIMInput()
        {
            Workbook bimWorkbook = null;
            bool workbookOpened = false;
            try
            {
                (bimWorkbook, workbookOpened) = ((FileTextBox)attributeDic["bimOutputFile_WC"]).OpenAndGetWorkbook(Globals.ThisAddIn.Application);
                designGroupsDic = new Dictionary<string, DesignGroupBIM>();
                ReadBimWallSheet(((AttributeTextBox)attributeDic["bimWallSheetName_WC"]).textBox.Text);
                ReadBimWallSheet(((AttributeTextBox)attributeDic["bimHsSheetName_WC"]).textBox.Text);
                ReadBimColSheet(((AttributeTextBox)attributeDic["bimColSheetName_WC"]).textBox.Text);

                void ReadBimWallSheet(string sheetName) 
                {
                    ExcelTableRange bimRange = new ExcelTableRange(sheetName, bimWorkbook, sheetName);
                    bimRange.GetUsedRangeFromEnd(1, 1, 1);
                    int startStoreyIndex = bimRange.GetHeaderIndex("DetailStartStorey");
                    int endStoreyIndex = bimRange.GetHeaderIndex("DetailEndStorey");
                    int mainBarIndex = bimRange.GetHeaderIndex("VerticalRebar");
                    int shearBarIndex = bimRange.GetHeaderIndex("HorizontalRebar");
                    int thicknessIndex = bimRange.GetHeaderIndex("Thickness");
                    int lengthIndex = -1;
                    try{ lengthIndex = bimRange.GetHeaderIndex("Length"); }
                    catch { }
                    

                    Range bimTableRange = bimRange.activeRange;
                    
                    string name = "";
                    foreach (Range row in bimTableRange.Rows)
                    {
                        string newName = row.Cells[1].Text;
                        if (newName != "") { name = newName; }
                        if (!designGroupsDic.ContainsKey(name)) { designGroupsDic[name] = new DesignGroupBIM(name, storeyMap); }
                        DesignGroupBIM wallRebar = designGroupsDic[name];
                        wallRebar.AddWallRow(row, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex, thicknessIndex, lengthIndex);
                    }
                }

                void ReadBimColSheet(string sheetName)
                {
                    ExcelTableRange bimRange = new ExcelTableRange(sheetName, bimWorkbook, sheetName);
                    bimRange.GetUsedRangeFromEnd(1, 1, 1);
                    int mainBarIndex = bimRange.GetHeaderIndex("MainRebar");
                    int shearBarIndex = bimRange.GetHeaderIndex("Stirrups");
                    int widthIndex = bimRange.GetHeaderIndex("Width");
                    int breathIndex = bimRange.GetHeaderIndex("Breadth");
                    int startStoreyIndex = bimRange.GetHeaderIndex("DetailStartStorey");
                    int endStoreyIndex = bimRange.GetHeaderIndex("DetailEndStorey");

                    Range bimTableRange = bimRange.activeRange;

                    string name = "";
                    foreach (Range row in bimTableRange.Rows)
                    {
                        string newName = row.Cells[1].Text;
                        if (newName != "") { name = newName; }
                        if (!designGroupsDic.ContainsKey(name)) { designGroupsDic[name] = new DesignGroupBIM(name, storeyMap); }
                        DesignGroupBIM designGroup = designGroupsDic[name];
                        designGroup.AddColRow(row, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex, widthIndex, breathIndex);
                    }
                }

                wallToGroupingDic = new Dictionary<string, string>();
                foreach (DesignGroupBIM designGroup in designGroupsDic.Values)
                {
                    designGroup.SortStories();
                    designGroup.MapIndividualPierLabels(wallToGroupingDic);
                }
            }
            catch (Exception ex) { throw new Exception("Error reading BIM Input table\n" + ex.Message); }
            finally
            {
                if (bimWorkbook != null && workbookOpened) { bimWorkbook.Close(); bimWorkbook = null; }
            }
        }
        #endregion

        #region Match
        private void MatchDesignLabels(Stopwatch totalStopwatch)
        {
            #region Init ETABS Array
            double[] verticalAsPrecReq = etabsRange.GetDataColumnAsDoubleArray("Required Reinf. Percentage");
            double[] shearRebarReq = etabsRange.GetDataColumnAsDoubleArray("Shear Rebar");
            double[] etabsThickness = etabsRange.GetDataColumnAsDoubleArray("Thickness");
            double[] etabsLength = etabsRange.GetDataColumnAsDoubleArray("Length");
            #endregion

            #region Init Write Arrays
            int numRows = etabsRange.activeRange.Rows.Count;
            string[] etabsStorey = GetContentsAsStringArray(etabsRange.GetDataColumnAsRange("Story"), false);
            string[] etabsPierLabel = GetContentsAsStringArray(etabsRange.GetDataColumnAsRange("Pier Label"), false);

            string[] matchedLabel = new string[numRows];
            string[] matchedDesignGroup = new string[numRows];
            
            string[] verticalBar = new string[numRows];
            object[] verticalAsProv = new object[numRows];
            object[] verticalAsPerc = new object[numRows];
            string[] verticalCheck  = new string[numRows];

            string[] horizontalBar = new string[numRows];
            object[] horizontalAsProv = new object[numRows];
            string[] horizontalCheck = new string[numRows];

            string[] thicknessCheck = new string[numRows];
            string[] lengthCheck = new string[numRows];
            #endregion

            #region Clear formats
            int outputColNum = ((AttributeTextBox)attributeDic["dispWallOuputColNum_WC"]).GetIntFromTextBox() - 1;
            Range verticalCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 4];
            Range horizontalCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 7];
            Range matchResultRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 10];
            Range thicknessCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 11];
            Range lengthCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 12];
            verticalCheckRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            horizontalCheckRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            matchResultRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            #endregion

            #region Error Functions
            Range errorFormatRange = null;
            HashSet<string> errorPiers = new HashSet<string>();

            void AddToErrorRange(Range targetRange, int rowNum)
            {
                if (errorFormatRange == null) { errorFormatRange = targetRange.Rows[rowNum + 1]; }
                else { errorFormatRange = errorFormatRange.Application.Union(errorFormatRange, targetRange.Rows[rowNum + 1]); }
                if (!errorPiers.Contains(etabsPierLabel[rowNum])) { errorPiers.Add(etabsPierLabel[rowNum]); }
            }
            #endregion

            for (int rowNum = 0; rowNum < numRows; rowNum++)
            {
                #region Match to Design Values
                try
                {
                    matchedLabel[rowNum] = labelMap.GetDesignName(etabsPierLabel[rowNum]);
                }
                catch (Exception ex)
                {
                    matchedDesignGroup[rowNum] = ex.Message;
                    AddToErrorRange(matchResultRange, rowNum);
                    continue;
                }


                if (wallToGroupingDic.ContainsKey(matchedLabel[rowNum])) 
                { 
                    matchedDesignGroup[rowNum] = wallToGroupingDic[matchedLabel[rowNum]]; 
                }
                else 
                { 
                    matchedDesignGroup[rowNum] = "Error finding design group";
                    AddToErrorRange(matchResultRange, rowNum);
                    continue; 
                }
                #endregion

                #region Calculate As
                RebarEntryBim entry = designGroupsDic[matchedDesignGroup[rowNum]].GetEntryFromEtabsStorey(etabsStorey[rowNum]);
                verticalBar[rowNum] = entry.vertcialBarString;
                (verticalAsProv[rowNum], verticalAsPerc[rowNum]) = entry.VerticalAs;
                if (double.IsNaN((double)verticalAsProv[rowNum])) { verticalAsProv[rowNum] = ""; }

                horizontalBar[rowNum] = entry.horizontalBarString;
                horizontalAsProv[rowNum] = entry.HorizontalAs;
                #endregion

                #region Check As
                if (verticalAsPrecReq[rowNum] < (double)verticalAsPerc[rowNum])
                {
                    verticalCheck[rowNum] = "Ok";
                }
                else
                {
                    verticalCheck[rowNum] = "Not Ok";
                    AddToErrorRange(verticalCheckRange, rowNum);
                }

                if (shearRebarReq[rowNum] < (double)horizontalAsProv[rowNum])
                {
                    horizontalCheck[rowNum] = "Ok";
                }
                else
                {
                    horizontalCheck[rowNum] = "Not Ok";
                    AddToErrorRange(horizontalCheckRange, rowNum);
                }
                #endregion

                #region Check Dimension
                double thickness = entry.Thickness;
                if (double.IsNaN(thickness))
                {
                    thicknessCheck[rowNum] = $"Error: Invalid thickness provided for check";
                    AddToErrorRange(thicknessCheckRange, rowNum);
                }
                else if (Math.Abs(etabsThickness[rowNum] - thickness) > 10)
                {
                    thicknessCheck[rowNum] = $"Error: ETABS thickness differs form length in BIM data. ETABS = {etabsThickness[rowNum]}, BIM = {thickness}";
                    AddToErrorRange(thicknessCheckRange, rowNum);
                }
                else
                {
                    thicknessCheck[rowNum] = $"Ok";
                }

                double length = entry.Length;
                if (double.IsNaN(length))
                {
                    lengthCheck[rowNum] = $"Warning: Invalid length provided for check";
                    AddToErrorRange(lengthCheckRange, rowNum);
                }
                else if ((etabsLength[rowNum] - length) > 10)
                {
                    lengthCheck[rowNum] = $"Error: ETABS length is greater than length in BIM data. ETABS = {etabsLength[rowNum]}, BIM = {length}";
                    AddToErrorRange(lengthCheckRange, rowNum);
                }
                else
                {
                    lengthCheck[rowNum] = $"Ok";
                }

                #endregion
            }

            WriteToExcelRangeAsCol(etabsRange.activeRange, 0, outputColNum, false, matchedLabel, 
                verticalBar, verticalAsProv, verticalAsPerc, verticalCheck, 
                horizontalBar, horizontalAsProv, horizontalCheck);
            WriteToExcelRangeAsCol(etabsRange.activeRange, 0 , outputColNum + 10, false, matchedDesignGroup, thicknessCheck, lengthCheck);
            
            #region Error handling
            if (errorFormatRange != null) { errorFormatRange.Font.Color = Color.Red; }
            string errorMsg = "Error encountered in the following etabs piers, please check result?\n";
            string errorMsgToClipboard = "";
            if (errorPiers.Count != 0)
            {
                foreach (string errorPier in errorPiers)
                {
                    errorMsg += errorPier + ", ";
                    errorMsgToClipboard += errorPier + "\n";
                }
                errorMsg = errorMsg.Substring(0, errorMsg.Length - 2); // remove last ", "
                errorMsgToClipboard = errorMsgToClipboard.Substring(0, errorMsgToClipboard.Length - 1); // remove last "\n"
                
                totalStopwatch.Stop();

                MessageBox.Show(errorMsg, "Warning");
                if (copyFailedCheck.Checked)
                {
                    Clipboard.SetText(errorMsgToClipboard);
                }
            }
            else { totalStopwatch.Stop(); }
            #endregion
        }

        private void MatchColDesignLabels(Stopwatch totalStopwatch)
        {
            #region Init ETABS Array
            string[] verticalAsPrecReqString = etabsRange.GetDataColumnAsStringArray("PMM Ratio or Rebar %");
            double[] shearRebarReqMaj = etabsRange.GetDataColumnAsDoubleArray("At Major");
            double[] shearRebarReqMin = etabsRange.GetDataColumnAsDoubleArray("At Minor");
            string[] sectionNames = etabsRange.GetDataColumnAsStringArray("Section");
            
            double ConvertVertPrec(string asReqString) {
                string stringValue = asReqString.Split(' ')[0];
                bool pass = double.TryParse(stringValue, out double verticalAsPrecReqDouble);
                if (pass) { return verticalAsPrecReqDouble; }
                else { throw new Exception("$Unable to parse {asReqString} into number"); }
            }
            #endregion

            #region Init Write Arrays
            int numRows = etabsRange.activeRange.Rows.Count;
            string[] etabsStorey = GetContentsAsStringArray(etabsRange.GetDataColumnAsRange("Story"), false);
            string[] etabsColLabel = GetContentsAsStringArray(etabsRange.GetDataColumnAsRange("Label"), false);

            string[] matchedLabel = new string[numRows];
            string[] matchedDesignGroup = new string[numRows];

            double[] width = new double[numRows];
            double[] breadth= new double[numRows];

            string[] verticalBar = new string[numRows];
            object[] verticalAsProv = new object[numRows];
            object[] verticalAsPerc = new object[numRows];
            string[] verticalCheck = new string[numRows];

            string[] horizontalBar = new string[numRows];
            object[] horizontalAsProvMaj = new object[numRows];
            object[] horizontalAsProvMin = new object[numRows];
            string[] horizontalCheck = new string[numRows];

            string[] thicknessCheck = new string[numRows];
            string[] lengthCheck = new string[numRows];
            #endregion

            #region Clear formats
            
            // These ranges are later used for reference to union merging formats so we're sticking with this approach I guess
            int outputColNum = ((AttributeTextBox)attributeDic["dispColOuputColNum_WC"]).GetIntFromTextBox() - 1;
            Range verticalCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 6];
            verticalCheckRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

            Range horizontalCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 10];
            horizontalCheckRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

            Range matchResultRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 13];
            matchResultRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

            Range thicknessCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 14];
            thicknessCheckRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

            Range lengthCheckRange = etabsRange.activeRange.Columns[1].Offset[0, outputColNum + 15];
            lengthCheckRange.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            #endregion

            #region Error Functions
            Range errorFormatRange = null;
            HashSet<string> errorPiers = new HashSet<string>();
            
            void AddToErrorRange(Range targetRange, int rowNum)
            {
                if (errorFormatRange == null) { errorFormatRange = targetRange.Rows[rowNum + 1]; }
                else { errorFormatRange = errorFormatRange.Application.Union(errorFormatRange, targetRange.Rows[rowNum + 1]); }
                if (!errorPiers.Contains(etabsColLabel[rowNum])) { errorPiers.Add(etabsColLabel[rowNum]); }
            }
            #endregion


            for (int rowNum = 0; rowNum < numRows; rowNum++)
            {
                #region Match to Design Values
                try
                {
                    matchedLabel[rowNum] = labelMap.GetDesignName(etabsColLabel[rowNum]);
                }
                catch (Exception ex)
                {
                    matchedDesignGroup[rowNum] = ex.Message;
                    AddToErrorRange(matchResultRange, rowNum);
                    continue;
                }
                

                if (wallToGroupingDic.ContainsKey(matchedLabel[rowNum]))
                {
                    matchedDesignGroup[rowNum] = wallToGroupingDic[matchedLabel[rowNum]];
                }
                else
                {
                    matchedDesignGroup[rowNum] = "Error finding design group";
                    AddToErrorRange(matchResultRange, rowNum);
                    continue;
                }
                #endregion
                
                #region Calculate As
                RebarEntryBim tryEntry = designGroupsDic[matchedDesignGroup[rowNum]].GetEntryFromEtabsStorey(etabsStorey[rowNum]);
                if (!(tryEntry is ColumnRebarEntryBim))
                {
                    matchedDesignGroup[rowNum] = $"Design group {matchedDesignGroup[rowNum]} is not a column type";
                    AddToErrorRange(matchResultRange, rowNum);
                    continue;
                }
                ColumnRebarEntryBim entry = (ColumnRebarEntryBim)tryEntry;

                verticalBar[rowNum] = entry.vertcialBarString;
                (verticalAsProv[rowNum], verticalAsPerc[rowNum]) = entry.VerticalAs;
                if (double.IsNaN((double)verticalAsProv[rowNum])) { verticalAsProv[rowNum] = ""; }

                horizontalBar[rowNum] = entry.horizontalBarString;
                horizontalAsProvMaj[rowNum] = entry.HorizontalAsMaj;
                horizontalAsProvMin[rowNum] = entry.HorizontalAsMin;
                #endregion

                #region Check As
                double verticalAsPrecReq = ConvertVertPrec(verticalAsPrecReqString[rowNum]);
                if (verticalAsPrecReq < (double)verticalAsPerc[rowNum])
                {
                    verticalCheck[rowNum] = "Ok";
                }
                else
                {
                    verticalCheck[rowNum] = "Not Ok";
                    AddToErrorRange(verticalCheckRange, rowNum);
                }
                
                if ((shearRebarReqMaj[rowNum] < (double)horizontalAsProvMaj[rowNum]) & (shearRebarReqMin[rowNum] < (double)horizontalAsProvMin[rowNum])) 
                {
                    horizontalCheck[rowNum] = "Ok";
                }
                else
                {
                    horizontalCheck[rowNum] = "Not Ok";
                    AddToErrorRange(horizontalCheckRange, rowNum);
                }
                #endregion

                #region Check Dimension

                #region Split Section Name
                (double etabsThickness, double etabsLength) = ConvertSectionToDouble(sectionNames[rowNum]);
                if (double.IsNaN(etabsThickness)) { continue; }
                (double thk, double len) ConvertSectionToDouble(string sectionName)
                {
                    Match match = Regex.Match(sectionName, @"^[A-Za-z]+(\d+)x(\d+)");
                    if (match.Success)
                    {
                        double thk = double.Parse(match.Groups[1].Value);
                        double len = double.Parse(match.Groups[2].Value);
                        return (thk, len);
                    }
                    else
                    {
                        Console.WriteLine("No match found.");
                        thicknessCheck[rowNum] = $"Error: Unable to split ETABS section name {sectionName}";
                        AddToErrorRange(thicknessCheckRange, rowNum);
                        lengthCheck[rowNum] = $"Error: Unable to split ETABS section name {sectionName}";
                        AddToErrorRange(lengthCheckRange, rowNum);
                        return (double.NaN, double.NaN);
                    }
                }
                #endregion

                double thickness = entry.Thickness;
                
                if (double.IsNaN(thickness))
                {
                    thicknessCheck[rowNum] = $"Error: Invalid thickness provided for check";

                    AddToErrorRange(thicknessCheckRange, rowNum);
                }
                else if (Math.Abs(etabsThickness - thickness) > 10)
                {
                    thicknessCheck[rowNum] = $"Error: ETABS thickness differs form length in BIM data. ETABS = {etabsThickness}, BIM = {thickness}";
                    AddToErrorRange(thicknessCheckRange, rowNum);
                }
                else
                {
                    thicknessCheck[rowNum] = $"Ok";
                }

                double length = entry.Length;
                if (double.IsNaN(length))
                {
                    lengthCheck[rowNum] = $"Warning: Invalid length provided for check";
                    AddToErrorRange(lengthCheckRange, rowNum);
                }
                else if ((etabsLength - length) > 10)
                {
                    lengthCheck[rowNum] = $"Error: ETABS length is greater than length in BIM data. ETABS = {etabsLength}, BIM = {length}";
                    AddToErrorRange(lengthCheckRange, rowNum);
                }
                else
                {
                    lengthCheck[rowNum] = $"Ok";
                }
                width[rowNum] = thickness;
                breadth[rowNum] = length;
                #endregion

            }
            WriteToExcelRangeAsCol(etabsRange.activeRange, 0, outputColNum, false, matchedLabel, width, breadth, 
                verticalBar, verticalAsProv, verticalAsPerc, verticalCheck, 
                horizontalBar, horizontalAsProvMaj, horizontalAsProvMin, horizontalCheck);
            WriteToExcelRangeAsCol(etabsRange.activeRange, 0, outputColNum + 13, false, matchedDesignGroup, thicknessCheck, lengthCheck);

            #region Error handling
            if (errorFormatRange != null) { errorFormatRange.Font.Color = Color.Red; }
            string errorMsg = "Error encountered in the following etabs columns, please check result.\n";
            string errorMsgToClipboard = "";
            if (errorPiers.Count != 0)
            {
                foreach (string errorPier in errorPiers)
                {
                    errorMsg += errorPier + ", ";
                    errorMsgToClipboard += errorPier + "\n";
                }
                errorMsg = errorMsg.Substring(0, errorMsg.Length - 2); // remove last ", "
                errorMsgToClipboard = errorMsgToClipboard.Substring(0, errorMsgToClipboard.Length - 1); // remove last "\n"

                totalStopwatch.Stop();

                MessageBox.Show(errorMsg, "Warning");
                if (copyFailedCheck.Checked)
                {
                    Clipboard.SetText(errorMsgToClipboard);
                }
            }
            else { totalStopwatch.Stop(); }
            #endregion
        }


        #endregion
    }
        public class ExcelTableRange
    {
        #region Init
        public Workbook workbook;
        public Worksheet worksheet;
        public string name;
        public ExcelTableRange(string name,string workbookPath, string worksheetNm)
        {
            workbook = OpenAndGetWorkbook(Globals.ThisAddIn.Application, workbookPath);
            try
            {
                worksheet = workbook.Worksheets[worksheetNm];
            }
            catch (Exception ex) { throw new ArgumentException($"Unable to find worksheet \"{worksheetNm}\" in workbook \"{workbook.Name}\"\n{ex.Message}"); }
        }
        public ExcelTableRange(string name, Workbook workbook, string worksheetNm)
        {
            this.workbook = workbook;
            this.name = name;
            try
            {
                worksheet = workbook.Worksheets[worksheetNm];
            }
            catch (Exception ex) { throw new ArgumentException($"Unable to find worksheet \"{worksheetNm}\" in workbook \"{workbook.Name}\"\n{ex.Message}"); }
        }
        public ExcelTableRange(string name, Worksheet worksheet)
        {
            this.name = name;
            this.worksheet = worksheet;
            workbook = worksheet.Parent;
        }
        #endregion

        #region Define Range
        public Range activeRange = null;
        public Range headerRange = null;
        public Range GetUsedRangeFromEnd(int headerRowNum, int colNum, int dataRowOffsetFromHeader = 0)
        {
            // Set dataRowOffsetFromHeader to 0 if there is no header
            Range lastRowCell = GetLastCellFromEnd(worksheet, colNum);
            Range lastColCell = GetLastCellFromEnd(worksheet, headerRowNum, XlDirection.xlToLeft);
            activeRange = worksheet.Range[lastRowCell, lastColCell.Offset[dataRowOffsetFromHeader]];
            if (dataRowOffsetFromHeader != 0) 
            { 
                headerRange = activeRange.Rows[1].Offset[-dataRowOffsetFromHeader]; 
                MapHeaderRange(); 
            }
            return activeRange;
        }
        public Range GetUsedRangeFromStart(int headerRowNum, int colNum, int dataRowOffsetFromHeader = 0)
        {
            // Set dataRowOffsetFromHeader to 0 if there is no header
            int rowNum = headerRowNum + dataRowOffsetFromHeader;
            Range startCell = worksheet.Cells[rowNum, colNum];
            Range endCell = GetLastCellFromStartCell(worksheet, headerRowNum, colNum);
            activeRange = worksheet.Range[startCell, endCell];
            
            if (dataRowOffsetFromHeader != 0)
            {
                headerRange = activeRange.Rows[1].Offset[-dataRowOffsetFromHeader];
                MapHeaderRange();
            }
            return activeRange;
        }
        public object[,] rangeContents = null;
        public object[,] GetDataAs2DObject(bool saveContentsToObject = true)
        {
            if (activeRange == null) { throw new Exception($"Data range is not defined for {name}"); }
            if (saveContentsToObject)
            {
                rangeContents = GetContentsAsObject2DArray(activeRange);
                return rangeContents;
            }
            else { return GetContentsAsObject2DArray(activeRange); }
        }
        public void SetActiveRange(Range range, int dataRowOffsetFromHeader)
        {
            activeRange = range;
            if (dataRowOffsetFromHeader != 0)
            {
                headerRange = activeRange.Rows[1].Offset[-dataRowOffsetFromHeader];
                MapHeaderRange();
            }
        }
        #endregion

        #region Headers
        public Range ReturnColumnRangeFromHeaderText(string headerText)
        {
            if (headerMapping.Count == 0) { MapHeaderRange(); }
            if (!headerMapping.ContainsKey(headerText)) { throw new Exception($"Header text {headerText} not found in header range {headerRange.Address[false, false]}"); }
            int colIndex = headerMapping[headerText];
            return activeRange.Columns[colIndex];
        }

        public Dictionary<string,int> headerMapping = new Dictionary<string,int>(); // colNum 1 indexed
        public void MapHeaderRange()
        {
            // colNum 1 indexed
            if (headerRange == null) { throw new Exception($"No header range defined for {name}"); }
            int colNum = 1;
            string address = headerRange.Address;
            foreach (Range cell in headerRange.Cells)
            {
                if (cell.Value2 == null) { continue; }
                string headerValue = cell.Value2.ToString().Trim();
                headerMapping.Add(headerValue, colNum);
                colNum++;
            }
        }

        public int GetHeaderIndex(string headerValue)
        {
            if (headerMapping.ContainsKey(headerValue))
            {
                return headerMapping[headerValue];
            }
            throw new Exception($"Header value \"{headerValue}\" not found in headers for \"{name}\"");
        }
        #endregion

        #region Data Columns
        Dictionary<string, int> dataColumnDic = new Dictionary<string, int>();
        public object[] GetDataColumnAsObject(string headerText)
        {
            Range colRange = GetDataColumnAsRange(headerText);
            return GetContentsAsObject1DArray(colRange);
        }
        public double[] GetDataColumnAsDoubleArray(string headerText)
        {
            Range colRange = GetDataColumnAsRange(headerText);
            return GetContentsAsDoubleArray(colRange,double.NaN);
        }
        public Range GetDataColumnAsRange(string headerText)
        {
            if (!headerMapping.ContainsKey(headerText)) { throw new Exception($"Header Text \"{headerText}\" not found in data \"{name}\""); }
            int colIndex = headerMapping[headerText];
            return activeRange.Columns[colIndex];
        }
        public string[] GetDataColumnAsStringArray(string headerText)
        {
            Range colRange = GetDataColumnAsRange(headerText);
            return GetContentsAsStringArray(colRange, false);
        }
        #endregion
    }

    public class EtabsToDesignMap
    {
        // This function has been generalised to map anything ETABS - DESIGN
        string name;
        Dictionary<string, string> etabsToDesignDic= new Dictionary<string, string>();
        Dictionary<string, string> designToEtabsDic = new Dictionary<string, string>();
        Dictionary<int, string> indexToEtabsDic = new Dictionary<int, string>();
        Dictionary<string, int> etabsToIndexDic = new Dictionary<string, int>();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="storeyTable"></param>
        /// <param name="mappingOptions">bool[etabsToDesign, designToEtabs, order]</param>
        public EtabsToDesignMap(string name, Range storeyTable, bool[] mappingOptions)
        {
            this.name = name;
            GetStoreyTable(storeyTable, mappingOptions);
        }

        private void GetStoreyTable(Range storeyTable, bool[] mappingOptions)
        {
            int index = storeyTable.Rows.Count;
            foreach (Range range in storeyTable.Rows)
            {
                string etabsStoreyName = range.Cells[1].Value2.ToString();
                string designStoreyName = range.Cells[2].Value2.ToString();

                if (mappingOptions[0]) {
                    if (etabsToDesignDic.ContainsKey(etabsStoreyName)) { throw new Exception($"Unable to add {etabsStoreyName} as there is a duplicate copy for mapping of {name} Table"); }
                    etabsToDesignDic.Add(etabsStoreyName, designStoreyName); 
                }
                if (mappingOptions[1]) {
                    if (designToEtabsDic.ContainsKey(designStoreyName)) { throw new Exception($"Unable to add {designStoreyName} as there is a duplicate copy for mapping of {name} Table"); }
                    designToEtabsDic.Add(designStoreyName, etabsStoreyName); 
                }
                if (mappingOptions[2])
                {
                    indexToEtabsDic.Add(index, etabsStoreyName);
                    etabsToIndexDic.Add(etabsStoreyName, index);
                    index--;
                }
            }
        }
        
        #region Get Values
        public string GetDesignName(string etabsStoreyName)
        {
            if (etabsToDesignDic.ContainsKey(etabsStoreyName)) { return etabsToDesignDic[etabsStoreyName]; }
            else { throw new ArgumentException($"Unable to find {etabsStoreyName} in {name} mapping table"); }
        }
        public string GetEtabsName(string designStoreyName)
        {
            if (designToEtabsDic.ContainsKey(designStoreyName)) { return designToEtabsDic[designStoreyName]; }
            else { throw new ArgumentException($"Unable to find {designStoreyName} in {name} mapping table"); }
        }
        public string GetEtabsName(int storeyIndex)
        {
            if (indexToEtabsDic.ContainsKey(storeyIndex))
            {
                return indexToEtabsDic[storeyIndex];
            }
            else
            {
                throw new ArgumentException($"Storey Index \"{storeyIndex}\" not found.");
            }
        }

        public string GetDesignName(int storeyIndex)
        {
            string etabsName = GetEtabsName(storeyIndex);
            return GetDesignName(etabsName);
        }

        public int GetStoreyIndex(string storeyName, string type)
        {
            if (etabsToIndexDic.Count == 0) { throw new Exception($"No indexing for mapping found for {name}"); }

            #region Convert to Design Storey Name
            string etabsStoreyName;
            if (type == "etabs")
            {
                etabsStoreyName = storeyName;
            }
            else if (type == "design")
            {
                if (designToEtabsDic.ContainsKey(storeyName))
                {
                    etabsStoreyName = designToEtabsDic[storeyName];
                }
                else
                {
                    throw new ArgumentException($"ETABS name \"{storeyName}\" not found.");
                }
            }
            else { throw new Exception($"Invalid type \"{type}\" provided"); }
            #endregion

            #region Get Index
            if (etabsToIndexDic.ContainsKey(etabsStoreyName))
            {
                return etabsToIndexDic[etabsStoreyName];
            }
            else
            {
                throw new ArgumentException($"Design Storey name \"{etabsStoreyName}\" not found.");
            }
            #endregion
        }
        #endregion

    }

    #region Assigned Rebar
    public class DesignGroupBIM
    {
        // Contains rebar information for a single pier label in the rebar table
        #region Init
        public string pierLabels;
        public Dictionary<int, RebarEntryBim> tableContents = new Dictionary<int, RebarEntryBim>(); // initial data in dictionary format, storey index points to data
        public List<string[]> tableContentsList = new List<string[]>(); // initial data in table format

        EtabsToDesignMap storeyMap;
        public DesignGroupBIM(string name, EtabsToDesignMap storeyMap)
        {
            // Used for matching rebars only
            pierLabels = name;
            this.storeyMap = storeyMap;
        }
        
        public void AddWallRow(Range row, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex, int thicknessIndex, int lengthIndex)
        {
            WallRebarEntryBim wallRebarEntry = new WallRebarEntryBim(row, storeyMap, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex, thicknessIndex, lengthIndex);
            tableContents.Add(wallRebarEntry.startStoreyNum, wallRebarEntry);
        }
        public void AddColRow(Range row, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex, int widthIndex, int breathIndex)
        {
            ColumnRebarEntryBim columnRebarEntry = new ColumnRebarEntryBim(row, storeyMap, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex, widthIndex, breathIndex);
            tableContents.Add(columnRebarEntry.startStoreyNum, columnRebarEntry);
        }

        #endregion

        #region Sort and Find Stories
        private int[] startStoreyNumSorted = null;
        private int[] endStoreyNumSorted = null;
        public void SortStories()
        {
            startStoreyNumSorted = tableContents.Keys.OrderBy(key => key).ToArray();

            endStoreyNumSorted = new int[startStoreyNumSorted.Length];
            
            for (int i = 0; i < startStoreyNumSorted.Length; i++)
            {
                RebarEntryBim wallRebarEntry = tableContents[startStoreyNumSorted[i]];
                endStoreyNumSorted[i] = wallRebarEntry.endStoreyNum;
            }
        }

        public RebarEntryBim GetEntryFromStoreyNum(int targetStoreyNum)
        {
            for (int i = 0; i < startStoreyNumSorted.Length; i++)
            {
                if (targetStoreyNum > startStoreyNumSorted[i] && targetStoreyNum <= endStoreyNumSorted[i])
                {
                    RebarEntryBim wallRebarEntry = tableContents[startStoreyNumSorted[i]];
                    return wallRebarEntry;
                }
            }

            return null;
            //throw new Exception("Unable to find target storey");
        }
        
        public RebarEntryBim GetEntryFromEtabsStorey(string etabsStoreyName)
        {
            int storeyNum = storeyMap.GetStoreyIndex(etabsStoreyName, "etabs");
            return GetEntryFromStoreyNum(storeyNum);
        }

        #endregion

        #region Map pier labels
        internal void MapIndividualPierLabels(Dictionary<string, string> wallToGroupingDic)
        {
            string[] parts = SplitAndTrim(pierLabels);
            foreach (string part in parts) 
            {
                if (wallToGroupingDic.ContainsKey(part)) { throw new Exception($"Error: Duplicate label \"{part}\" found\n" +
                    $"Pier Group 1: {wallToGroupingDic[part]}\n" +
                    $"Pier Group 2: {pierLabels}\n"); }
                wallToGroupingDic.Add(part, pierLabels);
            }
        }
        #endregion
    }

    public abstract class RebarEntryBim
    {
        #region Init
        public string vertcialBarString;
        public string horizontalBarString;
        public int startStoreyNum;
        public int endStoreyNum;

        public RebarEntryBim(Range row, EtabsToDesignMap storeyMap, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex)
        {
            string startStoreyName = row.Cells[startStoreyIndex].Value2.ToString();
            startStoreyNum = storeyMap.GetStoreyIndex(startStoreyName, "design");

            string endStoreyName = row.Cells[endStoreyIndex].Value2.ToString();
            endStoreyNum = storeyMap.GetStoreyIndex(endStoreyName, "design");
            vertcialBarString = row.Cells[mainBarIndex].Value2.ToString();
            horizontalBarString = row.Cells[shearBarIndex].Value2.ToString();
        }
        #endregion

        #region Calculated Values
        public abstract (double asProv, double asPerc) VerticalAs { get; }
        public abstract double HorizontalAs { get; }
        #endregion

        #region Dimension
        public abstract double Thickness { get; }
        public abstract double Length { get; }
        #endregion
    }

    public class WallRebarEntryBim: RebarEntryBim
    {
        // Contains rebar information for a single storey for a wall label in the rebar table
        #region Init
        public double thickness;
        public double length;
        /// <summary>
        /// 1 based indexes used
        /// </summary>
        public WallRebarEntryBim(Range row, EtabsToDesignMap storeyMap, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex, int thicknessIndex, int lengthIndex): base(row, storeyMap, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex)
        {
            //thickness = ReadDoubleFromCell(row.Cells[thicknessIndex]);
            //length = ReadDoubleFromCell(row.Cells[lengthIndex]);
            thickness = ReadDoubleFromCell2(row.Cells[thicknessIndex], double.NaN, double.NaN);
            if (lengthIndex != -1) { length = ReadDoubleFromCell2(row.Cells[lengthIndex], double.NaN, double.NaN); }
            else { length = double.NaN; }
        }
        #endregion

        #region Calculated values
        public override (double asProv, double asPerc) VerticalAs
        {
            get
            {
                (double dia, double spacing) = SplitRebarAndSpacing(vertcialBarString);
                double asProv = 2 * (Math.PI * Math.Pow(dia, 2)) / 4 * (1000 / spacing);

                double asPerc = asProv / thickness/1000 * 100;
                asProv = Math.Round(asProv, 0);
                asPerc = Math.Round(asPerc, 3);
                return (asProv, asPerc);
            }
        }

        public override double HorizontalAs
        {
            get
            {
                (double dia, double spacing) = SplitRebarAndSpacing(horizontalBarString);
                double asProv = 2 * (Math.PI * Math.Pow(dia, 2)) / 4 * (1000 / spacing);
                asProv = Math.Round(asProv, 0);
                return asProv;
            }
        }

        private (double dia, double spacing) SplitRebarAndSpacing(string inputString)
        {
            Regex regex = new Regex(@"H(\d+)-(\d+)");
            Match match = regex.Match(inputString);
            if (match.Success)
            {
                double dia = double.Parse(match.Groups[1].Value);
                double spacing = double.Parse(match.Groups[2].Value);
                return (dia, spacing);
            }
            else
            {
                throw new ArgumentException($"Unable to read \"{inputString}\" into format: Hxx-yyy");
            }
        }
        #endregion

        #region Dimensions
        public override double Thickness
        {
            get { return thickness; }
        }

        public override double Length
        {
            get { return length; }
        }
        #endregion
    }

    public class ColumnRebarEntryBim: RebarEntryBim
    {
        // Contains rebar information for a single storey for a column label in the rebar table
        #region Init
        public double width;
        public double breath;

        /// <summary>
        /// 1 based indexes used
        /// </summary>
        public ColumnRebarEntryBim(Range row, EtabsToDesignMap storeyMap, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex, int widthIndex, int breathIndex): base(row, storeyMap, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex)
        {
            width = ReadDoubleFromCell2(row.Cells[widthIndex], double.NaN, double.NaN);
            breath = ReadDoubleFromCell2(row.Cells[breathIndex], double.NaN, double.NaN);
        }
        #endregion

        #region Calculated values
        public override (double asProv, double asPerc) VerticalAs
        {
            get
            {
                (double number, double dia) = SplitRebarNumberAndDiameter(vertcialBarString);
                double asProv = (Math.PI * Math.Pow(dia, 2)) / 4 * number;
                
                double asPerc = asProv / (breath * width) * 100;
                asProv = Math.Round(asProv, 0);
                asPerc = Math.Round(asPerc, 3);
                return (double.NaN, asPerc); // return asProv as NaN since it should be hidden for columns
            }
        }

        public override double HorizontalAs // For pier check use
        {
            get
            {
                string[] parts = SplitAndTrim(horizontalBarString, '+');
                (double dia, double spacing, double num) = SplitRebarAndSpacingShear(parts[0]);
                double asProv = 2 * (Math.PI * Math.Pow(dia, 2)) / 4 * (1000 / spacing);
                asProv = Math.Round(asProv, 0);
                return asProv;
            }
        }
        public double HorizontalAsMaj // For col check use
        {
            get
            {
                string[] parts = SplitAndTrim(horizontalBarString, '+');
                (double dia, double spacing, double num) = SplitRebarAndSpacingShear(parts[0]);
                double asProv = num * (Math.PI * Math.Pow(dia, 2)) / 4 * (1000 / spacing);
                asProv = Math.Round(asProv, 0);
                return asProv;
            }
        }
        public double HorizontalAsMin // For col check use
        {
            get
            {
                string[] parts = SplitAndTrim(horizontalBarString, '+');
                (double dia, double spacing, double num) = SplitRebarAndSpacingShear(parts[1]);
                double asProv = num * (Math.PI * Math.Pow(dia, 2)) / 4 * (1000 / spacing);
                asProv = Math.Round(asProv, 0);
                return asProv;
            }
        }
        private (int number, double dia) SplitRebarNumberAndDiameter(string inputString)
        {
            Regex regex = new Regex(@"(\d+)H(\d+)");
            Match match = regex.Match(inputString);
            if (match.Success)
            {
                int number = Int32.Parse(match.Groups[1].Value);
                double dia = double.Parse(match.Groups[2].Value);
                return (number, dia);
            }
            else
            {
                throw new ArgumentException($"Unable to read \"{inputString}\" into format: xxHyy");
            }
        }
        private (double dia, double spacing, double num) SplitRebarAndSpacingShear(string inputString)
        {
            if (inputString[0] == 'H')
            {
                Regex regex = new Regex(@"H(\d+)-(\d+)");
                Match match = regex.Match(inputString);
                if (match.Success)
                {
                    double dia = double.Parse(match.Groups[1].Value);
                    double spacing = double.Parse(match.Groups[2].Value);
                    return (dia, spacing, 1);
                }
                else
                {
                    throw new ArgumentException($"Unable to read \"{inputString}\" into format: Hxx-yyy");
                }
            }
            else
            {
                Regex regex = new Regex(@"(\d+)H(\d+)-(\d+)");
                Match match = regex.Match(inputString);
                if (match.Success)
                {
                    double num = double.Parse(match.Groups[1].Value);
                    double dia = double.Parse(match.Groups[2].Value);
                    double spacing = double.Parse(match.Groups[3].Value);
                    return (dia, spacing, num);
                }
                else
                {
                    throw new ArgumentException($"Unable to read \"{inputString}\" into format: xHyy-zzz");
                }
            }
            
        }
        #endregion

        #region Dimensions
        public override double Thickness
        {
            get { return width; }
        }

        public override double Length
        {
            get { return breath; }
        }
        #endregion
    }

    #endregion
}




