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
using System.Windows.Forms.VisualStyles;
using System.Web;
using System.Text.RegularExpressions;

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
            var att = new CheckBoxAttribute("copyFromEtabs_WC", copyFromEtabsCheck);
            attTB = new FileTextBox("bimOutputFile_WC", dispBimOutputFile, setBimOutputFile);
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

        private void runRefCode_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch();
            try
            {
                stopwatch.Start();
                CreateTablePier();
                stopwatch.Stop();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            MessageBox.Show($"Completed, executiong time: {stopwatch.ElapsedTicks}", "Completed");
        }

        #region ChatGPT Code
        public void CreateTablePier()
        {
            // Declare and initialize variables for Excel workbooks and worksheets
            Application excelApp = Globals.ThisAddIn.Application;
            Workbook wallReport = null, sampleExcel = null, sampleData = null;
            Worksheet columnSheet = null, wallSheet = null, storySheet = null, labelSheet = null, pierDgn = null;

            // Paths to the workbooks
            string wallReportPath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
            string directoryPath = Path.GetDirectoryName(wallReportPath);
            string sampleExcelPath = Path.Combine(directoryPath, "Sample_EXCEL from BIM.xlsx");
            string sampleDataPath = Path.Combine(directoryPath, "Sample_Data Mapping.xlsx");
            //string wallReportPath = @"C:\Users\horei\Downloads\Intern 2024\Intern\Task 3\Wall Report.xlsm";
            //string sampleExcelPath = @"C:\Users\horei\Downloads\Intern 2024\Intern\Task 3\20240605 to Reiko_Task 2\Sample_EXCEL from BIM.xlsx";
            //string sampleDataPath = @"C:\Users\horei\Downloads\Intern 2024\Intern\Task 3\20240605 to Reiko_Task 2\Sample_Data Mapping.xlsx";

            try
            {
                // Open workbooks if they are not already open
                wallReport = OpenWorkbookIfNotOpen(excelApp, wallReportPath);
                sampleExcel = OpenWorkbookIfNotOpen(excelApp, sampleExcelPath);
                sampleData = OpenWorkbookIfNotOpen(excelApp, sampleDataPath);

                // Ensure workbooks are opened successfully
                if (wallReport == null || sampleExcel == null || sampleData == null)
                {
                    MessageBox.Show("One or more workbooks could not be opened.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Assign worksheets
                columnSheet = sampleExcel.Sheets["RC COLUMN"];
                wallSheet = sampleExcel.Sheets["RC WALL"];
                storySheet = sampleData.Sheets["Storey Mapping"];
                labelSheet = sampleData.Sheets["Label Mapping"];
                pierDgn = wallReport.Sheets["Pier Dgn Sum - Eurocode 2-2004"];

                if (columnSheet == null || wallSheet == null || storySheet == null || labelSheet == null || pierDgn == null)
                {
                    MessageBox.Show("One or more worksheets could not be found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Find the last rows
                int pdsLastRow = pierDgn.Cells[pierDgn.Rows.Count, 1].End(XlDirection.xlUp).Row;
                int rcwLastRow = wallSheet.Cells[wallSheet.Rows.Count, 1].End(XlDirection.xlUp).Row;
                int rccLastRow = columnSheet.Cells[columnSheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

                // Clear old data
                ClearRange(columnSheet, "R:V");
                ClearRange(wallSheet, "R:V");
                ClearRange(pierDgn, "Z4:AJ1048576");

                // Remove fill
                Range pierRange = pierDgn.Range["Z4:AJ1048576"];
                pierRange.Interior.Pattern = XlPattern.xlPatternNone;

                // Loop through PierDgn Range to perform data mapping
                foreach (Range cell in pierDgn.Range["B4:B" + pdsLastRow])
                {
                    string markKeyword = cell.Value;
                    Range foundCell = labelSheet.Columns["A"].Find(markKeyword, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                    if (foundCell != null)
                    {
                        pierDgn.Cells[cell.Row, 26].Value = labelSheet.Cells[foundCell.Row, 2].Value; // Column Z
                    }
                }

                // Find maximum story number
                int maxStory = 0;
                foreach (Range cell in pierDgn.Range["A4:A" + pdsLastRow])
                {
                    if (cell.Value != null && cell.Value.ToString().StartsWith("Story"))
                    {
                        string storeyName = cell.Value.ToString();
                        storeyName = storeyName.Substring(5);
                        string storeyNameWithoutChar = "";
                        foreach (char c in storeyName)
                        {
                            if (!char.IsNumber(c)) { break; }
                            storeyNameWithoutChar += c;
                        }

                        int storyNumber = int.Parse(storeyNameWithoutChar);
                        if (storyNumber > maxStory)
                        {
                            maxStory = storyNumber;
                        }
                    }
                }

                // Copy ranges from RC WALL and RC COLUMN
                wallSheet.Range["A1:C" + rcwLastRow].Copy(wallSheet.Range["R1"]);
                wallSheet.Range["G1:H" + rcwLastRow].Copy(wallSheet.Range["U1"]);
                columnSheet.Range["A1:C" + rccLastRow].Copy(columnSheet.Range["Q1"]);
                columnSheet.Range["F1:G" + rccLastRow].Copy(columnSheet.Range["T1"]);
                columnSheet.Range["I1:J" + rccLastRow].Copy(columnSheet.Range["V1"]);
                pierDgn.Range["A4:A" + pdsLastRow].Copy(pierDgn.Range["AJ4"]);

                // Update story labels to numbers
                UpdateStoryLabels(wallSheet, "S2:T" + rcwLastRow, maxStory);
                UpdateStoryLabels(columnSheet, "R2:S" + rccLastRow, maxStory);

                // Increment values in specific columns
                IncrementRangeValues(wallSheet, "S2:S" + rcwLastRow);
                IncrementRangeValues(columnSheet, "R2:R" + rccLastRow);

                // Process PierDgn for bar calculations
                ProcessPierDgnBars(pierDgn, pdsLastRow, maxStory);

                MessageBox.Show("Pier table creation completed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up COM objects to avoid memory leaks
                ReleaseObject(columnSheet);
                ReleaseObject(wallSheet);
                ReleaseObject(storySheet);
                ReleaseObject(labelSheet);
                ReleaseObject(pierDgn);
                ReleaseObject(wallReport);
                ReleaseObject(sampleExcel);
                ReleaseObject(sampleData);
            }
        }

        // Helper methods
        private Workbook OpenWorkbookIfNotOpen(Application excelApp, string path)
        {
            // Come back and refractor this to check if open workbook has the correct path
            string workbookName = Path.GetFileName(path);
            Workbook workbook; ;
            try
            {
                workbook = excelApp.Workbooks[workbookName];
            }
            catch { workbook = excelApp.Workbooks.Open(path); }
            return workbook;
        }

        private void ClearRange(Worksheet sheet, string range)
        {
            sheet.Range[range].Clear();
        }

        private void IncrementRangeValues(Worksheet sheet, string range)
        {
            foreach (Range cell in sheet.Range[range])
            {
                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double value))
                {
                    cell.Value = cell.Value + 1;
                }
            }
        }

        private void UpdateStoryLabels(Worksheet sheet, string range, int maxStory)
        {
            foreach (Range cell in sheet.Range[range])
            {
                if (cell.Value != null)
                {
                    string cellValue = cell.Value.ToString();
                    if (int.TryParse(cellValue, out int numericValue))
                    {
                        cell.Value = numericValue;
                    }
                    else
                    {
                        switch (cellValue)
                        {
                            case "Main Roof": cell.Value = maxStory + 1; break;
                            case "Mid Roof": cell.Value = maxStory + 2; break;
                            case "Upper Roof": cell.Value = maxStory + 3; break;
                            case "Foundation": cell.Value = 0; break;
                        }
                    }
                }
            }
        }

        private void ProcessPierDgnBars(Worksheet pierDgn, int lastRow, int maxStory)
        {
            // Example for calculating and processing bar data
            for (int i = 4; i <= lastRow; i++)
            {
                Range cell = pierDgn.Cells[i, "AA"];
                if (cell.Value != null)
                {
                    string cellValue = cell.Value.ToString();
                    // Implement your bar calculation logic here
                }
            }
        }

        private void ReleaseObject(object obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
        #endregion

        #region Check Walls
        private void checkWalls_Click(object sender, EventArgs e)
        {
            //List<TrackedRange> trackedRanges = null;
            try
            {
                Stopwatch totalStopwatch = Stopwatch.StartNew();
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                ReadMapping();
                ReadETABSInput();
                ReadBIMInput();
                MatchDesignLabels();

                //object[] designLabels = MatchDesignLabels();
                //MatchRebars();

                //Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                //pierLabelRange.Worksheet.Activate();

                totalStopwatch.Stop();
                MessageBox.Show($"Total Execution Time: {totalStopwatch.ElapsedMilliseconds} ms", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                //HighlightChangesForMatchRebar(trackedRanges);
                //#region Release Dictionaries
                //rebarDic = null;
                //storeyTracker = null;
                //#endregion
                storeyMap = null;
                labelMap = null;
                etabsRange = null;
                wallToGroupingDic = null;
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
        private void ReadETABSInput()
        {
            try
            {
                if (copyFromEtabsCheck.Checked)
                {
                    #region Clear current sheet
                    Range destinationRange = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range["A4"];
                    Range clearRange = destinationRange.Worksheet.Range["A4:AJ1048576"];
                    clearRange.Clear();
                    #endregion

                    #region Copy
                    Workbook etabsWorkbook = ((FileTextBox)attributeDic["etabsOutputFile_WC"]).OpenAndGetWorkbook(Globals.ThisAddIn.Application);
                    string worksheetNm = ((AttributeTextBox)attributeDic["etabsSheetName_WC"]).textBox.Text;
                    etabsRange = new ExcelTableRange("etabs", etabsWorkbook, worksheetNm);
                    etabsRange.GetUsedRangeFromEnd(2, 1, 2);
                    Range copyRange = etabsRange.activeRange;
                    string copyRangeAddress = copyRange.Address;
                    copyRange.Copy(destinationRange);
                    etabsWorkbook.Close();
                    #endregion

                    etabsRange = new ExcelTableRange("etabs", destinationRange.Worksheet);
                    etabsRange.SetActiveRange(destinationRange.Worksheet.Range[copyRangeAddress], 2);
                }
                else
                {
                    Range destinationRange = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range["A4"];
                    etabsRange = new ExcelTableRange("etabs", destinationRange.Worksheet);
                    etabsRange.GetUsedRangeFromStart(2, 1, 2);
                }
            }
            catch (Exception ex) { throw new Exception("Error reading ETABS Input table\n" + ex.Message); }
        }        
        Dictionary<string, AssignedWallRebarBIM> rebarAssignmentDic;
        Dictionary<string, string> wallToGroupingDic;
        private void ReadBIMInput()
        {
            Workbook bimWorkbook = null;
            try
            {
                bimWorkbook = ((FileTextBox)attributeDic["bimOutputFile_WC"]).OpenAndGetWorkbook(Globals.ThisAddIn.Application);
                rebarAssignmentDic = new Dictionary<string, AssignedWallRebarBIM>();
                ReadBimWallSheet(((AttributeTextBox)attributeDic["bimWallSheetName_WC"]).textBox.Text);
                ReadBimWallSheet(((AttributeTextBox)attributeDic["bimHsSheetName_WC"]).textBox.Text);
                
                void ReadBimWallSheet(string sheetName) 
                {
                    //string sheetName = ((AttributeTextBox)attributeDic["bimWallSheetName_WC"]).textBox.Text;
                    ExcelTableRange bimRange = new ExcelTableRange("bim", bimWorkbook, sheetName);
                    bimRange.GetUsedRangeFromEnd(1, 1, 1);
                    int mainBarIndex = bimRange.headerMapping["VerticalRebar"];
                    int shearBarIndex = bimRange.headerMapping["HorizontalRebar"];
                    int thicknessIndex = bimRange.headerMapping["Thickness"];
                    int startStoreyIndex = bimRange.headerMapping["DetailStartStorey"];
                    int endStoreyIndex = bimRange.headerMapping["DetailEndStorey"];

                    Range bimTableRange = bimRange.activeRange;
                    
                    string name = "";
                    foreach (Range row in bimTableRange.Rows)
                    {
                        string newName = row.Cells[1].Text;
                        if (newName != "") { name = newName; }
                        if (!rebarAssignmentDic.ContainsKey(name)) { rebarAssignmentDic[name] = new AssignedWallRebarBIM(name, storeyMap); }
                        AssignedWallRebarBIM wallRebar = rebarAssignmentDic[name];
                        wallRebar.AddRow(row, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex, thicknessIndex);
                    }
                }

                wallToGroupingDic = new Dictionary<string, string>();
                foreach (AssignedWallRebarBIM wallRebar in rebarAssignmentDic.Values)
                {
                    wallRebar.SortStories();
                    wallRebar.MapIndividualPierLabels(wallToGroupingDic);
                }
            }
            catch (Exception ex) { throw new Exception("Error reading BIM Input table\n" + ex.Message); }
            finally
            {
                if (bimWorkbook != null) { bimWorkbook.Close(); }
            }
        }
        #endregion

        #region Match
        private void MatchDesignLabels()
        {
            #region Init ETABS Array
            double[] verticalAsReq = etabsRange.GetDataColumnAsDoubleArray("Required Reinf. Percentage");
            double[] shearRebarReq = etabsRange.GetDataColumnAsDoubleArray("Shear Rebar");
            #endregion

            #region Init Write Arrays
            int numRows = etabsRange.activeRange.Rows.Count;
            string[] etabsStorey = GetContentsAsStringArray(etabsRange.GetDataColumnAsRange("Story"), false);
            string[] etabsPierLabel = GetContentsAsStringArray(etabsRange.GetDataColumnAsRange("Pier Label"), false);
            //string[] matchedStorey = new string[numRows];
            string[] matchedLabel = new string[numRows];
            string[] matchedDesignGroup = new string[numRows];
            string[] verticalBar = new string[numRows];
            double[] verticalAsProv = new double[numRows];
            double[] verticalAsPerc = new double[numRows];
            string[] verticalCheck  = new string[numRows];

            string[] horizontalBar = new string[numRows];
            double[] horizontalAsProv = new double[numRows];
            string[] horizontalCheck = new string[numRows];
            #endregion


            for (int rowNum = 0; rowNum < numRows; rowNum++)
            {
                #region Match to Design Values
                //matchedStorey[rowNum] = storeyMap.GetDesignName(etabsStorey[rowNum]);
                matchedLabel[rowNum] = labelMap.GetDesignName(etabsPierLabel[rowNum]);
                
                if (wallToGroupingDic.ContainsKey(matchedLabel[rowNum])) 
                { 
                    matchedDesignGroup[rowNum] = wallToGroupingDic[matchedLabel[rowNum]]; 
                }
                else { matchedDesignGroup[rowNum] = "Error finding design group"; continue; }
                #endregion

                #region Calculate As
                WallRebarEntryBim entry = rebarAssignmentDic[matchedDesignGroup[rowNum]].GetEntryFromEtabsStorey(etabsStorey[rowNum]);
                verticalBar[rowNum] = entry.vertcialBarString;
                (verticalAsProv[rowNum], verticalAsPerc[rowNum]) = entry.VerticalAs;

                horizontalBar[rowNum] = entry.horizontalBarString;
                horizontalAsProv[rowNum] = entry.HorizontalAs;
                #endregion

                #region Check As
                if (verticalAsReq[rowNum] < verticalAsPerc[rowNum])
                {
                    verticalCheck[rowNum] = "Ok";
                }
                else
                {
                    verticalCheck[rowNum] = "Not Ok";
                }

                if (shearRebarReq[rowNum] < horizontalAsProv[rowNum])
                {
                    horizontalCheck[rowNum] = "Ok";
                }
                else
                {
                    horizontalCheck[rowNum] = "Not Ok";
                }
                #endregion

            }

            WriteToExcelRangeAsCol(etabsRange.activeRange, 0, 25, false, matchedLabel, verticalBar, verticalAsProv, verticalAsPerc, verticalCheck, horizontalBar, horizontalAsProv, horizontalCheck);
            WriteToExcelRangeAsCol(etabsRange.activeRange, 0 , 35, false, matchedDesignGroup);
            
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
                headerMapping.Add(cell.Value2.ToString(), colNum);
                colNum++;
            }
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
        #endregion
    }

    public class EtabsToDesignMap
    {
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

                if (mappingOptions[0]) { etabsToDesignDic.Add(etabsStoreyName, designStoreyName); }
                if (mappingOptions[1]) { designToEtabsDic.Add(designStoreyName, etabsStoreyName); }
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


    #region Wall Rebar
    public class AssignedWallRebarBIM
    {
        // Contains rebar information for a single pier label in the rebar table
        #region Init
        public string pierLabels;
        public Dictionary<int, WallRebarEntryBim> tableContents = new Dictionary<int, WallRebarEntryBim>(); // initial data in dictionary format, storey index points to data
        public List<string[]> tableContentsList = new List<string[]>(); // initial data in table format

        EtabsToDesignMap storeyMap;
        public AssignedWallRebarBIM(string name, EtabsToDesignMap storeyMap)
        {
            // Used for matching rebars only
            pierLabels = name;
            this.storeyMap = storeyMap;
        }
        
        public void AddRow(Range row, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex, int thicknessIndex)
        {
            WallRebarEntryBim wallRebarEntry = new WallRebarEntryBim(row, storeyMap, startStoreyIndex, endStoreyIndex, mainBarIndex, shearBarIndex, thicknessIndex);
            tableContents.Add(wallRebarEntry.startStoreyNum, wallRebarEntry);
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
                WallRebarEntryBim wallRebarEntry = tableContents[startStoreyNumSorted[i]];
                endStoreyNumSorted[i] = wallRebarEntry.endStoreyNum;
            }
        }

        public WallRebarEntryBim GetEntryFromStoreyNum(int targetStoreyNum)
        {
            for (int i = 0; i < startStoreyNumSorted.Length; i++)
            {
                if (targetStoreyNum > startStoreyNumSorted[i] && targetStoreyNum <= endStoreyNumSorted[i])
                {
                    WallRebarEntryBim wallRebarEntry = tableContents[startStoreyNumSorted[i]];
                    return wallRebarEntry;
                }
            }

            return null;
            //throw new Exception("Unable to find target storey");
        }
        public WallRebarEntryBim GetEntryFromEtabsStorey(string etabsStoreyName)
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
                if (wallToGroupingDic.ContainsKey(part)) { throw new Exception($"Error: Duplicate pier label \"{part}\" found\n" +
                    $"Pier Group 1: {wallToGroupingDic[part]}\n" +
                    $"Pier Group 2: {pierLabels}\n"); }
                wallToGroupingDic.Add(part, pierLabels);
            }
        }
        #endregion
    }

    public class WallRebarEntryBim
    {
        // Contains rebar information for a single storey for a pier label in the rebar table
        #region Init
        public string vertcialBarString;
        public string horizontalBarString;
        public int startStoreyNum;
        public int endStoreyNum;
        public double thickness;

        /// <summary>
        /// 1 based indexes used
        /// </summary>
        /// <param name="row"></param>
        /// <param name="storeyMap"></param>
        /// <param name="mainBarIndex"></param>
        /// <param name="shearBarIndex"></param>
        public WallRebarEntryBim(Range row, EtabsToDesignMap storeyMap, int startStoreyIndex, int endStoreyIndex, int mainBarIndex, int shearBarIndex, int thicknessIndex)
        {
            string startStoreyName = row.Cells[startStoreyIndex].Value2.ToString();
            startStoreyNum = storeyMap.GetStoreyIndex(startStoreyName, "design");

            string endStoreyName = row.Cells[endStoreyIndex].Value2.ToString();
            endStoreyNum = storeyMap.GetStoreyIndex(endStoreyName, "design");
            vertcialBarString = row.Cells[mainBarIndex].Value2.ToString();
            horizontalBarString = row.Cells[shearBarIndex].Value2.ToString();

            // Thickness
            Range cell = row.Cells[thicknessIndex];
            if (cell.Value2 is double) { thickness = cell.Value2; }
            else {
                bool canParse = double.TryParse(cell.Value2.ToString(), out thickness);
                if (!canParse) { throw new ArgumentException($"Unable to parse value {cell.Value2} at cell {cell.Worksheet.Name}!{cell.Address[false, false]} into number."); }
            }
        }
        #endregion

        #region Calculated values
        public (double asProv, double asPerc) VerticalAs
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

        public double HorizontalAs
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
                throw new ArgumentException($"No match found for {inputString} at ");
            }
            
        }
        #endregion
    }
    
    //public class WallRebarEntry

    #endregion
}


