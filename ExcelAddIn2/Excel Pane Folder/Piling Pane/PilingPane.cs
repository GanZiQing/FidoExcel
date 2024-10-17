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

using System.Runtime.InteropServices;
//using Autodesk.AutoCAD.Interop;
//using Autodesk.AutoCAD.Interop.Common;
using static ExcelAddIn2.CommonUtilities;
using ExcelAddIn2.Piling;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class PilingPane : UserControl
    {
        #region Initialisers
        Workbook ThisWorkBook;
        Microsoft.Office.Interop.Excel.Application ThisApplication;
        DocumentProperties AllCustProps;
        Dictionary<string, AttributeTextBox> TextBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();

        private void InitializeBeaver()
        {
            string folderPath = Path.GetDirectoryName(ThisApplication.ActiveWorkbook.FullName);
            string fileName = Path.GetFileNameWithoutExtension(ThisApplication.ActiveWorkbook.FullName) + "_ErrorLog.txt";
            Beaver.Initialize(folderPath, fileName);
        }

        public PilingPane()
        {
            InitializeComponent();
            ThisApplication = Globals.ThisAddIn.Application;
            ThisWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            AllCustProps = ThisWorkBook.CustomDocumentProperties;
            CreateAttributes();
        }

        private void CreateAttributes()
        {
            #region Draw BH
            RangeTextBox drawRange_pile = new RangeTextBox("drawRange_pile", DispDrawRange, SetDrawRange, "row");
            TextBoxAttributeDic.Add("drawRange_pile", drawRange_pile);

            RangeTextBox rockRange_pile = new RangeTextBox("rockRange_pile", DispRockRange, SetRockRange, "range");
            TextBoxAttributeDic.Add("rockRange_pile", rockRange_pile);

            RangeTextBox spt100Range_pile = new RangeTextBox("spt100Range_pile", dispSpt100Range, setSpt100Range, "range");
            TextBoxAttributeDic.Add("spt100Range_pile", spt100Range_pile);

            RangeTextBox notRockRange_pile = new RangeTextBox("notRockRange_pile", DispNotRockRange, SetNotRockRange, "range");
            TextBoxAttributeDic.Add("notRockRange_pile", notRockRange_pile);
            #endregion

            #region Copy Soil Data
            // Input
            RangeTextBox inputSoil_pile = new RangeTextBox("inputSoil_pile", dispSoilInputData, setSoilInputData, "range");
            TextBoxAttributeDic.Add("inputSoil_pile", inputSoil_pile);

            RangeTextBox inputRockTypes_pile = new RangeTextBox("inputRockTypes_pile", dispRockTypeInput, setRockTypeInput, "range");
            TextBoxAttributeDic.Add("inputRockTypes_pile", inputRockTypes_pile);

            RangeTextBox inputNsfTypes_pile = new RangeTextBox("inputNsfTypes_pile", dispNsfTypeInput, setNsfTypeInput, "range");
            TextBoxAttributeDic.Add("inputNsfTypes_pile", inputNsfTypes_pile);

            // Output
            SheetTextBox refSheet_pile = new SheetTextBox("refSheet_pile", dispRefSheet, setRefSheet);
            TextBoxAttributeDic.Add("refSheet_pile", refSheet_pile);

            RangeTextBox bhRLCell_pile = new RangeTextBox("bhRLCell_pile", dispBhRlCell, setBhRlCell, "cell", false);
            TextBoxAttributeDic.Add("bhRLCell_pile", bhRLCell_pile);

            RangeTextBox soilDest_pile = new RangeTextBox("soilDest_pile", dispSoilDest, setSoilDest, "range", false);
            TextBoxAttributeDic.Add("soilDest_pile", soilDest_pile);

            RangeTextBox fsRange_pile = new RangeTextBox("fsRange_pile", dispFsRange, setFsRange, "range", false);
            TextBoxAttributeDic.Add("fsRange_pile", fsRange_pile);

            RangeTextBox qbRange_pile = new RangeTextBox("qbRange_pile", dispQbRange, setQbRange, "row", false);
            TextBoxAttributeDic.Add("qbRange_pile", qbRange_pile);

            RangeTextBox rockStart_pile = new RangeTextBox("rockStart_pile", dispRockStart, setRockStart, "cell", false);
            TextBoxAttributeDic.Add("rockStart_pile", rockStart_pile);

            RangeTextBox spt100Start_pile = new RangeTextBox("spt100Start_pile", dispSpt100Start, setSpt100Start, "cell", false);
            TextBoxAttributeDic.Add("spt100Start_pile", spt100Start_pile);

            RangeTextBox effRange_pile = new RangeTextBox("effRange_pile", dispEffRange, setEffRange, "range", false);
            TextBoxAttributeDic.Add("effRange_pile", effRange_pile);

            // Sheet
            AttributeTextBox appendName_pile = new AttributeTextBox("appendName_pile ", dispAppendName, true);
            appendName_pile.type = "filename";
            TextBoxAttributeDic.Add("appendName_pile ", appendName_pile);
            #endregion


            #region Design Piles
            MultipleSheetsAttribute sheetsToRun_pile = new MultipleSheetsAttribute("sheetsToRun_pile", setSheetsToRun);
            OtherAttributeDic.Add("sheetsToRun_pile", sheetsToRun_pile);

            MultipleSheetsAttribute KeepSheets = new MultipleSheetsAttribute("KeepSheets", delSheets, true);
            OtherAttributeDic.Add("KeepSheets", KeepSheets);

            AttributeTextBox effLower_pile = new AttributeTextBox("effLower_pile", dispEfficiencyLower, true);
            effLower_pile.type = "double";
            effLower_pile.defaultValue = "0.98";
            effLower_pile.RefreshTextBox();
            TextBoxAttributeDic.Add("effLower_pile", effLower_pile);

            AttributeTextBox effUpper_pile = new AttributeTextBox("effUpper_pile", dispEfficiencyUpper, true);
            effUpper_pile.type = "double";
            effUpper_pile.defaultValue = "1.0";
            effUpper_pile.RefreshTextBox();
            TextBoxAttributeDic.Add("effUpper_pile", effUpper_pile);
            #endregion
        }

        public TabPage GetPageTaskPane(int tabNum)
        {
            TabControl.TabPageCollection MyTabPages = pilingTabControl.TabPages;
            TabPage ThisTabPage = MyTabPages[tabNum];
            return ThisTabPage;
        }
        #endregion

        #region Draw BH
        //static AcadApplication acApp;
        //static AcadDocument acDoc;
        //static AcadModelSpace modelSpace;
        //static Application excelApp;
        //static Workbook activeWB;


        static double baseLevel;
        static double offsetX;
        static double offsetY;
        static double lateralSpacing;
        static double offsetHead;

        static double width;
        static double textHeight;
        //static double width = 1.5;
        //static double textHeight = 0.65;

        private void DrawBH_Click(object sender, EventArgs e)
        {
            //ProgressHelper.RunWithProgress((worker, progressTrackerLocal) => RunFunction(worker, progressTrackerLocal));

            //void RunFunction(BackgroundWorker worker, ProgressTracker progressTrackerLocal)
            //{
            //    #region Create Logger
            //    progressTrackerLocal.UpdateStatus("Initialising");
            //    InitializeBeaver();
            //    #endregion

            //    #region Get Apps
            //    try
            //    {
            //        (acApp, acDoc, modelSpace) = GetAcApp();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Unable to get autocad application" + ex.Message, "Error");
            //        return;
            //    }

            //    // Get Excel Object
            //    excelApp = Globals.ThisAddIn.Application;
            //    activeWB = excelApp.ActiveWorkbook;
            //    Range runRange = null;
            //    try
            //    {
            //        runRange = ((RangeTextBox)TextBoxAttributeDic["drawRange_pile"]).GetRangeFromFullAddress();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show($"Unable to read Draw Range.\n\n{ex.Message}", "Error");
            //        return;
            //    }
            //    Worksheet activeWS = runRange.Worksheet;
            //    #endregion

            //    #region Read Excel Info
            //    progressTrackerLocal.UpdateStatus("Reading BH Info");
            //    List<Borehole> bhInfo = null;
            //    // Excel parameters
            //    try
            //    {
            //        GetParams(activeWS);
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error reading parameters from excel sheet info\n\n" + ex.Message);
            //        return;
            //    }

            //    // Rock types
            //    HashSet<string> rockTypes;
            //    HashSet<string> notRockTypes;
            //    try
            //    {
            //        rockTypes = GetRockType(true);
            //        notRockTypes = GetRockType(false);
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error reading rock type from excel pane \n\n" + ex.Message);
            //        return;
            //    }

            //    if (rockTypes.Count == 0)
            //    {
            //        Beaver.LogError("Warning no rock type provided, dwg will not reflect this rock type");
            //    }
            //    if (notRockTypes.Count == 0)
            //    {
            //        Beaver.LogError("Warning no rock type to design as soil provided, dwg will not reflect this rock type");
            //    }

            //    // Borehole table
            //    try
            //    {
            //        bhInfo = GetBHInfo(runRange, rockTypes, notRockTypes, worker);
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error reading borehole info from excel table\n\n" + ex.Message);
            //        return;
            //    }

            //    if (worker.CancellationPending)
            //    {
            //        return;
            //    }
            //    #endregion

            //    #region Draw
            //    // Draw The reference line 

            //    try
            //    {
            //        if (checkDrawRef.Checked)
            //        {
            //            DrawReferenceLine(bhInfo);
            //        }

            //        int maxProgress = bhInfo.Count;
            //        int currentProgress = 0;
            //        foreach (Borehole bh in bhInfo)
            //        {
            //            if (worker.CancellationPending)
            //            {
            //                break;
            //            }
            //            DrawOneBH(bh);
            //            currentProgress += 1;
            //            progressTrackerLocal.UpdateStatus($"Drawing {bh.name}");
            //            worker.ReportProgress(ConvertToProgress(currentProgress, maxProgress));
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        acDoc.Regen(AcRegenType.acActiveViewport);
            //        MessageBox.Show("Unable to complete draw operation\n\n" + ex.Message, "Error");
            //        return;
            //    }
            //    #endregion

            //    acDoc.Regen(AcRegenType.acActiveViewport);
            //    Beaver.CheckLog();
            //    if (!worker.CancellationPending)
            //    {
            //        progressTrackerLocal.UpdateStatus("Completed");
            //        MessageBox.Show("Completed", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            //    }
            //}
        }
        #region Archive
        //private void DrawBH_Click(object sender, EventArgs e)
        //{
        //    ProgressTracker progressTracker = new ProgressTracker();

        //    ProgressHelper.RunWithProgress((worker, e2) =>
        //    {
        //        #region Create Logger
        //        progressTracker.UpdateStatus("Initialising");
        //        InitializeBeaver();
        //        #endregion

        //        #region Get Apps
        //        try
        //        {
        //            (acApp, acDoc, modelSpace) = GetAcApp();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Unable to get autocad application" + ex.Message, "Error");
        //            return;
        //        }

        //        // Get Excel Object
        //        excelApp = Globals.ThisAddIn.Application;
        //        activeWB = excelApp.ActiveWorkbook;
        //        Range runRange = null;
        //        try
        //        {
        //            runRange = ((RangeTextBox)RangeAttributeDic["drawRange_pile"]).GetRangeFromFullAddress();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show($"Unable to read Draw Range.\n\n{ex.Message}", "Error");
        //            return;
        //        }
        //        Worksheet activeWS = runRange.Worksheet;
        //        #endregion

        //        #region Read Excel Info
        //        progressTracker.UpdateStatus("Reading BH Info");
        //        List<Borehole> bhInfo = null;
        //        // Excel parameters
        //        try
        //        {
        //            GetParams(activeWS);
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Error reading parameters from excel sheet info\n\n" + ex.Message);
        //            return;
        //        }

        //        // Rock types
        //        HashSet<string> rockTypes;
        //        HashSet<string> notRockTypes;
        //        try
        //        {
        //            rockTypes = GetRockType(true);
        //            notRockTypes = GetRockType(false);
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Error reading rock type from excel pane \n\n" + ex.Message);
        //            return;
        //        }

        //        if (rockTypes.Count == 0)
        //        {
        //            Beaver.LogError("Warning no rock type provided, dwg will not reflect this rock type");
        //        }
        //        if (notRockTypes.Count == 0)
        //        {
        //            Beaver.LogError("Warning no rock type to design as soil provided, dwg will not reflect this rock type");
        //        }

        //        // Borehole table
        //        try
        //        {
        //            bhInfo = GetBHInfo(runRange, rockTypes, notRockTypes, worker);
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Error reading borehole info from excel table\n\n" + ex.Message);
        //            return;
        //        }
        //        #endregion

        //        #region Draw
        //        // Draw The reference line 

        //        try
        //        {
        //            if (checkDrawRef.Checked)
        //            {
        //                DrawReferenceLine(bhInfo);
        //            }

        //            int maxProgress = bhInfo.Count;
        //            int currentProgress = 0;
        //            foreach (Borehole bh in bhInfo)
        //            {
        //                if (worker.CancellationPending)
        //                {
        //                    e2.Cancel = true;
        //                    break;
        //                }
        //                DrawOneBH(bh);
        //                currentProgress += 1;
        //                progressTracker.UpdateStatus($"Drawing {bh.name}");
        //                worker.ReportProgress(ConvertToProgress(currentProgress, maxProgress));
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Unable to complete draw operation\n\n" + ex.Message, "Error");
        //        }
        //        #endregion

        //        acDoc.Regen(AcRegenType.acActiveViewport);
        //        Beaver.CheckLog();
        //        MessageBox.Show("Completed", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

        //    }, progressTracker);
        //}






        //private void DrawBH_OG(object sender, EventArgs e)
        //{
        //    #region Create Logger
        //    InitializeBeaver();
        //    #endregion

        //    #region Get Apps
        //    try
        //    {
        //        (acApp, acDoc, modelSpace) = GetAcApp();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Unable to get autocad application" + ex.Message, "Error");
        //        return;
        //    }

        //    // Get Excel Object
        //    excelApp = Globals.ThisAddIn.Application;
        //    activeWB = excelApp.ActiveWorkbook;
        //    Range runRange = null;
        //    try
        //    {
        //        runRange = ((RangeTextBox)RangeAttributeDic["drawRange_pile"]).GetRangeFromFullAddress();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Unable to read Draw Range.\n\n{ex.Message}","Error");
        //        return;
        //    }
        //    Worksheet activeWS = runRange.Worksheet;
        //    #endregion

        //    #region Read Excel Info
        //    List<Borehole> bhInfo = null;
        //    // Excel parameters
        //    try
        //    {
        //        GetParams(activeWS);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error reading parameters from excel sheet info\n\n" + ex.Message);
        //        return;
        //    }

        //    // Rock types
        //    HashSet<string> rockTypes;
        //    HashSet<string> notRockTypes;
        //    try
        //    {
        //        rockTypes = GetRockType(true);
        //        notRockTypes = GetRockType(false);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error reading rock type from excel pane \n\n" + ex.Message);
        //        return;
        //    }

        //    if (rockTypes.Count == 0)
        //    {
        //        Beaver.LogError("Warning no rock type provided, dwg will not reflect this rock type");
        //    }
        //    if (notRockTypes.Count == 0)
        //    {
        //        Beaver.LogError("Warning no rock type to design as soil provided, dwg will not reflect this rock type");
        //    }

        //    // Borehole table
        //    try
        //    {
        //        //bhInfo = GetBHInfo(runRange, rockTypes, notRockTypes);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error reading borehole info from excel table\n\n" + ex.Message);
        //        return;
        //    }


        //    #endregion

        //    #region Draw
        //    // Draw The reference line 

        //    try
        //    {
        //        if (checkDrawRef.Checked)
        //        {
        //            DrawReferenceLine(bhInfo);
        //        }

        //        foreach (Borehole bh in bhInfo)
        //        {
        //            DrawOneBH(bh);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Unable to complete draw operation\n\n" + ex.Message, "Error");
        //    }
        //    #endregion

        //    acDoc.Regen(AcRegenType.acActiveViewport);
        //    Beaver.CheckLog();
        //    MessageBox.Show("Completed", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

        //}
        #endregion

        #region Helpers
        #region Connections
        //static (AcadApplication, AcadDocument, AcadModelSpace) GetAcApp()
        //{
        //    try
        //    {
        //        AcadApplication acApp = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application.18");
        //        //acadApp = (AcadApplication)Marshal.GetActiveObject("acad.exe");
        //        //acadApp.Visible = true;
        //        AcadDocument acDoc = acApp.ActiveDocument;
        //        AcadModelSpace modelSpace = acDoc.ModelSpace;
        //        return (acApp, acDoc, modelSpace);
        //    }
        //    catch (Exception ex)
        //    {
        //        string msg = "Unable to connection to AutoCAD";
        //        Console.WriteLine($"{msg}: \n {ex.Message}");
        //        throw new Exception(msg);
        //    }
        //}
        #endregion

        #region Excel Info
        static void GetParams(Worksheet thisSheet)
        {
            //double baseLevel = ReadDoubleFromCell(thisSheet.Range["B1"]);
            //double offsetX = ReadDoubleFromCell(thisSheet.Range["E1"]);
            //double offsetY = ReadDoubleFromCell(thisSheet.Range["E2"]);
            //double lateralSpacing = ReadDoubleFromCell(thisSheet.Range["B2"]);
            //double offsetHead = ReadDoubleFromCell(thisSheet.Range["H1"]);

            baseLevel = ReadDoubleFromCell(thisSheet.Range["B1"]);
            lateralSpacing = ReadDoubleFromCell(thisSheet.Range["B2"]);
            width = ReadDoubleFromCell(thisSheet.Range["B3"]);
            textHeight = ReadDoubleFromCell(thisSheet.Range["B4"]);

            offsetX = ReadDoubleFromCell(thisSheet.Range["E1"]);
            offsetY = ReadDoubleFromCell(thisSheet.Range["E2"]);
            offsetHead = ReadDoubleFromCell(thisSheet.Range["E3"]);
            
            return;
        }

        private HashSet<string> GetRockType(bool getRock = true)
        {
            HashSet<string> rockTypes = new HashSet<string>();

            #region Check if text box is empty
            string type = null;
            if (getRock)
            {
                type = "rockRange_pile";
            }
            else
            {
                type = "notRockRange_pile";
            }
            RangeTextBox thisRockCategory = (RangeTextBox)TextBoxAttributeDic[type]; 
            if (thisRockCategory.textBox.Text == "")
            {
                return rockTypes;
            }
            #endregion

            #region Check if excel cell is empty
            Range rockRange = thisRockCategory.GetRangeFromFullAddress();
            if (rockRange.Value2 == null || rockRange.Value2 ==  "")
            {
                return rockTypes;
            }
            #endregion

            foreach (Range cell in rockRange)
            {
                string value = cell.Value2;
                value = value.Trim();
                if (value.Length > 0)
                {
                    rockTypes.Add(value);
                }
            }
            return rockTypes;
        }

        static List<Borehole> GetBHInfo(Range thisRange, HashSet<string> rockTypes, HashSet<string> notRockTypes, BackgroundWorker worker)
        {
            int numCol = thisRange.Columns.Count;
            int startRowNum = thisRange.Row;
            int startColNum = thisRange.Column;
            List<Borehole> boreholes = new List<Borehole>();
            int currentProgress = 0;
            int maxProgress = numCol;
            for (int locColNum = 0; locColNum < numCol; locColNum += 3)
            {
                if (worker.CancellationPending)
                {
                    break;
                }

                int colNum = startColNum + locColNum;
                Borehole thisBH = new Borehole(thisRange.Worksheet.Cells[startRowNum, colNum], rockTypes, notRockTypes);
                boreholes.Add(thisBH);
                currentProgress = locColNum;
                worker.ReportProgress(ConvertToProgress(currentProgress, maxProgress));
            }
            return boreholes;
        }
        #endregion

        #region Drawing Boreholes
        static void DrawReferenceLine(List<Borehole> bhInfo)
        {
            //#region Start coordinates
            //double[] startPoint = new double[3];
            //startPoint[0] = offsetX - width;
            //startPoint[1] = baseLevel + offsetY;
            //#endregion
            //#region End coordinates
            //double[] endPoint = new double[3];
            //foreach (Borehole bh in bhInfo)
            //{
            //    if ((bh.index - 1) * lateralSpacing > endPoint[1])
            //    {
            //        endPoint[0] = (bh.index - 1) * lateralSpacing + 3 * width;
            //    }
            //}
            //endPoint[0] += startPoint[0];
            //endPoint[1] = startPoint[1];
            //#endregion

            //acDoc.ModelSpace.AddLine(startPoint, endPoint);

            //#region Label
            //double[] insertPoint = OffsetPoint(startPoint, -0.5, 0);
            //string msg = $"Reference Level: {baseLevel}mSHD";
            //AcadMText label = modelSpace.AddMText(insertPoint,0,msg);
            //label.Height = textHeight;
            //label.AttachmentPoint = AcAttachmentPoint.acAttachmentPointMiddleRight;
            //insertPoint = OffsetPoint(insertPoint, label.Width, 0);
            //label.InsertionPoint = insertPoint;
            //label.Update();
            //#endregion
        }

        static void DrawOneBH(Borehole bh)
        {

            //#region Draw side and top lines
            //double[] startPoint = new double[3];
            //startPoint[0] = offsetX + (bh.index - 1) * lateralSpacing;
            //startPoint[1] = offsetY + bh.reducedLevel;
            //double[] endPoint = (double[])startPoint.Clone();
            //endPoint[1] -= bh.depth[bh.depth.Length - 1];
            //AcadLine leftLine = acDoc.ModelSpace.AddLine(startPoint, endPoint);

            //startPoint[0] += width;
            //endPoint[0] += width;
            //AcadLine rightLine = acDoc.ModelSpace.AddLine(startPoint, endPoint);

            //// Draw top line
            //endPoint = (double[])startPoint.Clone();
            //startPoint = OffsetPoint(endPoint, -width, 0);
            //AcadLine firstHorizontalLine = acDoc.ModelSpace.AddLine(startPoint, endPoint);
            //#endregion

            //#region Write BH Name and level
            //double[] insertPoint = OffsetPoint(startPoint, 0, 0);
            //insertPoint[1] = offsetY + offsetHead + baseLevel + textHeight * 3;
            //AcadMText text = acDoc.ModelSpace.AddMText(insertPoint, lateralSpacing, $"{bh.name}");
            //text.Height = textHeight;

            //insertPoint[1] -= textHeight * 1.5;
            //text = acDoc.ModelSpace.AddMText(insertPoint, lateralSpacing, $"{bh.reducedLevel}mSHD");
            //text.Height = textHeight;

            //#endregion

            //#region Draw each depth
            //double[] refPoint = (double[])startPoint.Clone();
            //for (int rowNum = 0; rowNum < bh.depth.Length; rowNum++)
            //{
            //    // Draw Horizontal Line
            //    startPoint = (double[])refPoint.Clone();
            //    startPoint[1] -= bh.depth[rowNum];
            //    endPoint = OffsetPoint(startPoint, width, 0);
            //    AcadLine bottomLine = acDoc.ModelSpace.AddLine(startPoint, endPoint);

            //    // Insert Text
            //    if (rowNum != bh.depth.Length - 1) // skip last row
            //    {
            //        insertPoint = OffsetPoint(endPoint, 0.5, textHeight / 2);
            //        string textString = $"SPT {bh.sptValue[rowNum]}, {bh.rockType[rowNum]}";
            //        text = acDoc.ModelSpace.AddMText(insertPoint, lateralSpacing - 2 * width, textString);
            //        text.Height = textHeight;
            //    }

            //    // Insert Hatch
            //    //if (rowNum != 0 && bh.sptValue[rowNum] == 100 && bh.sptValue[rowNum - 1] == 100)
            //    if (rowNum != bh.depth.Length-1 && bh.sptValue[rowNum] >= 100) // skip last row
            //    {
            //        //AcadPolyline rectangle = DrawRectangle(startPoint, width, (bh.depth[rowNum] - bh.depth[rowNum - 1]));
            //        AcadPolyline rectangle = DrawRectangle(startPoint, width, (bh.depth[rowNum] - bh.depth[rowNum + 1]));
            //        AcadEntity[] boundary = new AcadEntity[1];
            //        boundary[0] = (AcadEntity)rectangle;

            //        AcadHatch hatch;
            //        //if (bh.isRock[rowNum - 1] == 1 && bh.isRock[rowNum] == 1)
            //        if (bh.isRock[rowNum] == 1) // is rock
            //            {
            //            hatch = acDoc.ModelSpace.AddHatch(0, "EARTH", false);
            //            hatch.AppendOuterLoop(boundary);
            //            hatch.color = ACAD_COLOR.acBlue;
            //        }
            //        else if (bh.isRock[rowNum] == 0) // is overwritten rock
            //        {
            //            hatch = acDoc.ModelSpace.AddHatch(0, "EARTH", false);
            //            hatch.AppendOuterLoop(boundary);
            //            hatch.color = ACAD_COLOR.acYellow;
            //        }
            //        else // is SPT 100 soil
            //        {
            //            hatch = acDoc.ModelSpace.AddHatch(0, "EARTH", false);
            //            hatch.AppendOuterLoop(boundary);
            //            hatch.color = ACAD_COLOR.acCyan;
            //        }
            //        hatch.PatternScale = 0.1;
            //        rectangle.Delete();
            //    }
            //}
            //#endregion

            //#region Add BH Termination text
            //insertPoint = OffsetPoint(startPoint, 0, -textHeight * 1.5);
            //string msg = $"{bh.reducedLevel - bh.depth[bh.depth.Length - 1]}mSHD\n({bh.depth[bh.depth.Length - 1]}m)";
            //text = acDoc.ModelSpace.AddMText(insertPoint, lateralSpacing, msg);
            //text.Height = textHeight;
            //text.color = ACAD_COLOR.acRed;
            //#endregion
        }

        //public static AcadPolyline DrawRectangle(double[] origin, double dx, double dy)
        //{
        //    double[] point1 = origin;
        //    double[] point2 = OffsetPoint(point1, dx, 0);
        //    double[] point3 = OffsetPoint(point2, 0, dy);
        //    double[] point4 = OffsetPoint(point3, -dx, 0);
        //    double[] point = new double[3 * 5];
        //    point1.CopyTo(point, 0);
        //    point2.CopyTo(point, 3);
        //    point3.CopyTo(point, 6);
        //    point4.CopyTo(point, 9);
        //    point1.CopyTo(point, 12);

        //    AcadPolyline rectangle = modelSpace.AddPolyline(point);
        //    rectangle.Closed = true;
        //    return rectangle;
        //}
        #endregion

        #endregion

        #endregion

        #region AGS Parser

        [STAThread]
        private void importAGS_Click(object sender, EventArgs e)
        {
            try
            {
                #region User Input
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Select .ags File";
                openFileDialog.Filter = "AGS files (*.ags)|*.ags|Text files (*.txt)|*.txt|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Process terminated.", "Error");
                    return;
                }

                string filePath = openFileDialog.FileName;
                HashSet<string> rockTypes = GetContentsAsStringHash(((RangeTextBox)TextBoxAttributeDic["rockRange_pile"]).GetRangeFromFullAddress());
                HashSet<string> spt100Types = GetContentsAsStringHash(((RangeTextBox)TextBoxAttributeDic["spt100Range_pile"]).GetRangeFromFullAddress());
                #endregion

                InitializeBeaver();

                Dictionary<string, BoreholeAGS> bhDict = new Dictionary<string, BoreholeAGS>();

                #region Read AGS File
                try
                {
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        // Use a StreamReader to read from the stream with specified encoding
                        using (StreamReader streamReader = new StreamReader(fileStream))
                        {
                            string line;
                            while ((line = streamReader.ReadLine()) != null)
                            {
                                if (line.StartsWith("\"**GEOL\""))
                                {
                                    try
                                    {
                                        ReadGEOL(streamReader, ref bhDict);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception($"Error reading GEOL for line:\n" +
                                            $"line\n" + ex.Message);
                                    }
                                }
                                else if (line.StartsWith("\"**HOLE\""))
                                {
                                    try
                                    {
                                        ReadHOLE(streamReader, ref bhDict);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception($"Error reading HOLE for line:\n" +
                                            $"line\n" + ex.Message);
                                    }
                                }
                                else if (line.StartsWith("\"**ISPT\""))
                                {
                                    try
                                    {
                                        ReadISPT(streamReader, ref bhDict);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception($"Error reading ISPT for line:\n" +
                                            $"line\n" + ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred with reading AGS file:\n" + ex.Message);
                    Beaver.CheckLog();
                    return;
                }
                #endregion

                #region Post Processing and Print to Excel
                Range selRange = ThisApplication.Selection;

                try
                {
                    ThisApplication.ScreenUpdating = false;
                    bool printDescription = printDescriptionCheck.Checked;
                    bool skipEmptySPT = checkRemoveNoSPT.Checked;
                    bool fillSoilType = checkFillSoilType.Checked;
                    bool defaultSPT = checkDefaultSPT.Checked;
                    bool compressOutput = checkCompressOutput.Checked;
                    int index = 1;
                    foreach (BoreholeAGS bh in bhDict.Values)
                    {
                        bh.WriteBHToExcel(selRange, rockTypes, spt100Types, skipEmptySPT, fillSoilType, printDescription, defaultSPT, compressOutput);
                        selRange.Offset[2, 1].Value2 = index;
                        if (printDescription)
                        {
                            selRange = selRange.Offset[0, 4];
                        }
                        else
                        {
                            selRange = selRange.Offset[0, 3];
                        }
                        index++;
                    }
                }
                catch (Exception ex)
                {
                    ThisApplication.ScreenUpdating = true;
                    MessageBox.Show($"Unable to write to excel. \n\n{ex.Message}", "Error");
                    Beaver.CheckLog();
                    return;
                }
                finally
                {
                    ThisApplication.ScreenUpdating = true;
                }
                #endregion

                Beaver.CheckLog();
                MessageBox.Show($"Extracted values for {bhDict.Count} boreholes.", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                return;
            }
            
        }

        #region Read Category
        static void ReadGEOL(StreamReader streamReader, ref Dictionary<string, BoreholeAGS> bhDict)
        {
            #region Header 
            Dictionary<string, int?> headerColNum = new Dictionary<string, int?>(){
            { "HOLE_ID", null },
            { "GEOL_TOP", null },
            { "GEOL_BASE", null },
            { "GEOL_GEOL", null },
            { "GEOL_DESC", null },
            };

            GetHeaderColNum(streamReader, ref headerColNum, "**GEOL");
            #endregion

            #region Data
            string line;
            while ((line = streamReader.ReadLine()) != null)
            {
                if (line == "")
                {
                    break;
                }

                string[] lineArray = ParseLine(line);

                #region Get BH Object
                string bhName = lineArray[(int)headerColNum["HOLE_ID"]];
                if (!bhDict.ContainsKey(bhName))
                {
                    bhDict[bhName] = new BoreholeAGS(true);
                    bhDict[bhName].name = bhName;
                }
                BoreholeAGS thisBh = bhDict[bhName];
                #endregion

                #region Add BH info
                string inputTopDepth = "";
                string inputBotDepth = "";
                string inputSoilType = "";
                string inputSoilDescription = "";
                foreach (KeyValuePair<string, int?> entry in headerColNum)
                {
                    string headerName = entry.Key;
                    int colNum = (int)entry.Value;
                    switch (headerName)
                    {
                        case "GEOL_TOP":
                            inputTopDepth = lineArray[colNum];
                            break;
                        case "GEOL_BASE":
                            inputBotDepth = lineArray[colNum];
                            break;
                        case "GEOL_GEOL":
                            inputSoilType = lineArray[colNum];
                            break;
                        case "GEOL_DESC":
                            inputSoilDescription = lineArray[colNum];
                            break;
                    }
                }

                if (inputTopDepth == "" || inputSoilType == "")
                {
                    //ThrowExceptionBox($"Unable to parse line {line}");
                    //Beaver.LogError($"Unable to parse line {line}");
                }
                else
                {
                    thisBh.AddBhSoilTypeToList(inputTopDepth, inputBotDepth, inputSoilType, inputSoilDescription);
                }
                #endregion
            }
            #endregion
        }

        static void ReadHOLE(StreamReader streamReader, ref Dictionary<string, BoreholeAGS> bhDict)
        {
            #region Header 
            Dictionary<string, int?> headerColNum = new Dictionary<string, int?>(){
            { "HOLE_ID", null },
            { "HOLE_GL", null },
            };

            GetHeaderColNum(streamReader, ref headerColNum, "**HOLE");
            #endregion

            #region Data
            string line;
            while ((line = streamReader.ReadLine()) != null)
            {
                if (line == "")
                {
                    break;
                }

                string[] lineArray = ParseLine(line);

                #region Get BH Object
                string bhName = lineArray[(int)headerColNum["HOLE_ID"]];
                if (!bhDict.ContainsKey(bhName))
                {
                    bhDict[bhName] = new BoreholeAGS(true);
                    bhDict[bhName].name = bhName;
                }
                BoreholeAGS thisBh = bhDict[bhName];
                #endregion

                #region Add BH info
                string bhRL = "";
                foreach (KeyValuePair<string, int?> entry in headerColNum)
                {
                    string headerName = entry.Key;
                    int colNum = (int)entry.Value;
                    switch (headerName)
                    {
                        case "HOLE_GL":
                            bhRL = lineArray[colNum];
                            break;
                    }
                }

                if (bhRL == "")
                {
                    ThrowExceptionBox($"Unable to parse line {line}");
                }

                thisBh.AddBhRl(bhRL);
                #endregion
            }
            #endregion
        }

        static void ReadISPT(StreamReader streamReader, ref Dictionary<string, BoreholeAGS> bhDict)
        {
            #region Header 
            Dictionary<string, int?> headerColNum = new Dictionary<string, int?>(){
            { "HOLE_ID", null },
            { "ISPT_TOP", null },
            { "ISPT_MAIN", null },
            };
            GetHeaderColNum(streamReader, ref headerColNum, "**ISPT");
            #endregion

            #region Data
            string line;
            while ((line = streamReader.ReadLine()) != null)
            {
                if (line == "")
                {
                    break;
                }

                string[] lineArray = ParseLine(line);

                #region Get BH Object
                string bhName = lineArray[(int)headerColNum["HOLE_ID"]];
                if (!bhDict.ContainsKey(bhName))
                {
                    bhDict[bhName] = new BoreholeAGS(true);
                    bhDict[bhName].name = bhName;
                    MessageBox.Show($"Unable to find Borehole {bhName} for SPT input");
                    continue;
                }
                BoreholeAGS thisBh = bhDict[bhName];
                #endregion

                #region Add BH info
                string inputDepth = "";
                string inputSptValue = "";

                foreach (KeyValuePair<string, int?> entry in headerColNum)
                {
                    string headerName = entry.Key;
                    int colNum = (int)entry.Value;
                    switch (headerName)
                    {
                        case "ISPT_TOP":
                            inputDepth = lineArray[colNum];
                            break;
                        case "ISPT_MAIN":
                            inputSptValue = lineArray[colNum];
                            break;
                    }
                }

                if (inputDepth == "" || inputSptValue == "")
                {
                    ThrowExceptionBox($"Unable to parse line {line}");
                }

                thisBh.AddSPT(inputDepth, inputSptValue);
                #endregion
            }
            #endregion
        }

        static string[] ParseLine(string line)
        {
            string[] lineArray = line.Split(new string[] { "\",\"" }, StringSplitOptions.None);
            for (int i = 0; i < lineArray.Length; i++)
            {
                lineArray[i] = lineArray[i].Trim('\"');
                //lineArray[i] = lineArray[i].TrimStart('*');
            }
            return lineArray;
        }

        static void GetHeaderColNum(StreamReader streamReader, ref Dictionary<string, int?> headerColNum, string type)
        {
            string line;
            while (true)
            {
                line = streamReader.ReadLine(); // This skips one line upon break for units 
                if ((line).Trim('"')[0] != '*')
                {
                    break;
                }
                string[] headers = line.Split(',');
                for (int index = 0; index < headers.Length; index++)
                {
                    string header = headers[index].Trim('\"');
                    header = header.Trim('*');
                    if (headerColNum.ContainsKey(header))
                    {
                        headerColNum[header] = index;
                    }
                }
            }
            // Check all entry found.
            foreach (KeyValuePair<string, int?> entry in headerColNum)
            {
                if (entry.Value == null)
                {
                    ThrowExceptionBox($"Unable to find index for {entry.Key} for {type}");
                }
            }
        }

        #endregion

        #region Backup Write to Excel
        //private void WriteArrayToExcelRange(Range thisRange, int rowOff, int colOff, params Array[] arrays)
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

        //    // Initiate object
        //    object[,] dataArray = new object[numRow, numCol];
        //    for (int col = 0; col < arrays.Length; col++)
        //    {
        //        for (int row = 0; row < arrays[col].Length; row++)
        //        {
        //            dataArray[row, col] = arrays[col].GetValue(row);
        //        }
        //    }

        //    // Write to Excel
        //    thisRange.Application.ScreenUpdating = true;
        //    Range startCell = thisRange.Offset[rowOff, colOff];
        //    Range endCell = startCell.Offset[numRow - 1, numCol - 1];
        //    Range writeRange = thisRange.Worksheet.Range[startCell, endCell];
        //    writeRange.Value2 = dataArray;
        //    thisRange.Application.ScreenUpdating = true;
        //}

        //private void WriteListToExcelRange(Range thisRange, int rowOff, int colOff, params List<object>[] listOfLists)
        //{
        //    // This code takes any number of arrays (of various types) and outputs them into excel 
        //    // Output order depends on order of the input array
        //    // Output location is the first cell of the current selection, offset by rowOff and colOff

        //    // Find number of rows and columns
        //    int numRow = 0;
        //    int numCol = listOfLists.Count();
        //    foreach (List<object> thisList in listOfLists)
        //    {
        //        if (thisList.Count() > numRow)
        //        {
        //            numRow = thisList.Count(); // Finds max number of rows out of all the various arrays
        //        }
        //    }

        //    // Initiate object
        //    object[,] dataArray = new object[numRow, numCol];
        //    int col = 0;
        //    foreach (List<object> thisList in listOfLists)
        //    {
        //        int row = 0;
        //        foreach (object entry in thisList)
        //        {
        //            dataArray[row, col] = entry.ToString();
        //            row += 1;
        //        }
        //        col += 1;
        //    }

        //    // Write to Excel
        //    thisRange.Application.ScreenUpdating = true;
        //    Range startCell = thisRange.Offset[rowOff, colOff];
        //    Range endCell = startCell.Offset[numRow - 1, numCol - 1];
        //    Range writeRange = thisRange.Worksheet.Range[startCell, endCell];
        //    writeRange.Value2 = dataArray;
        //    thisRange.Application.ScreenUpdating = true;
        //}
        #endregion

        #endregion

        #region Raw AGS
        private void basicAGS_Click(object sender, EventArgs e)
        {
            #region User Input
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select .ags File";
            openFileDialog.Filter = "AGS files (*.ags)|*.ags|Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Process terminated.", "Error");
                return;
            }
            string filePath = openFileDialog.FileName;
            #endregion

            try
            {
                ThisApplication.ScreenUpdating = false;
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Use a StreamReader to read from the stream with specified encoding
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        string line;
                        while ((line = streamReader.ReadLine()) != null)
                        {
                            if (line.StartsWith("\"**"))
                            {
                                string sheetTitle = line.Trim('\"');
                                sheetTitle = sheetTitle.TrimStart('*');
                                Worksheet newSheet = ThisApplication.Worksheets.Add(After: ThisWorkBook.Worksheets[ThisWorkBook.Worksheets.Count]);
                                newSheet.Name = sheetTitle;

                                int rowNum = 6;
                                int colNum = 1;

                                while ((line = streamReader.ReadLine()) != null)
                                {
                                    colNum = 1;
                                    string[] parts = ParseLine(line);

                                    foreach (string part in parts)
                                    {
                                        string value = part;
                                        if ((value.StartsWith("*")))
                                        {
                                            value = value.TrimStart('*');
                                        }

                                        Range targetCell = newSheet.Cells[colNum][rowNum];
                                        targetCell.Value2 = value;
                                        colNum++;
                                    }
                                    rowNum++;

                                    if (line == "")
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred with reading AGS file:\n" + ex.Message);
                Beaver.CheckLog();
                return;
            }
            finally
            {
                ThisApplication.ScreenUpdating = true;
            }
        }
        #endregion

        #region Clean Up AGS
        private void removeCont_Click(object sender, EventArgs e)
        {
            #region User Input
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select .ags File";
            openFileDialog.Filter = "AGS files (*.ags)|*.ags|Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Process terminated.", "Error");
                return;
            }
            string filePath = openFileDialog.FileName;
            #endregion

            #region Output File
            string outputFileName = Path.GetFileNameWithoutExtension(filePath) + "_cleaned.ags";
            string outputPath = Path.Combine(Path.GetDirectoryName(filePath), outputFileName);
            #endregion

            InitializeBeaver();

            #region Read AGS File
            try
            {
                using (StreamWriter writer = new StreamWriter(outputPath))
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Use a StreamReader to read from the stream with specified encoding
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        string firstLine = streamReader.ReadLine();
                        string nextLine = streamReader.ReadLine();
                        while (nextLine != null)
                        {
                            if (nextLine.StartsWith("\"<CONT>\""))
                            {
                                CombineLines(streamReader, ref firstLine, ref nextLine);
                            }
                            writer.WriteLine(firstLine);
                            firstLine = nextLine;
                            nextLine = streamReader.ReadLine();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred with reading AGS file:\n" + ex.Message);
                Beaver.CheckLog();
                return;
            }
            #endregion

            Beaver.CheckLog();
            MessageBox.Show($"Finish cleaning file.", "Completed");
        }

        private void CombineLines(StreamReader streamReader, ref string firstLine, ref string nextLine)
        {
            // Runs through lines to concat line with subsequent <Cont> lines until no more
            string[] finalLineParts = ParseLine(firstLine);

            while (nextLine.StartsWith("\"<CONT>\""))
            {
                string[] nextLineParts = ParseLine(nextLine);
                for (int i = 1; i < nextLineParts.Length; i++)
                {
                    if (nextLineParts[i] != "")
                    {
                        finalLineParts[i] += nextLineParts[i];
                    }
                }
                nextLine = streamReader.ReadLine();
            }

            string finalLine = "\"";
            finalLine += String.Join("\",\"", finalLineParts);
            finalLine += "\"";

            firstLine = finalLine;
        }
        #endregion

        #region Copy Soil Data
        private void copySoilData_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                try
                {
                    #region Main
                    ThisApplication.ScreenUpdating = false;
                    #region Read BH to Run
                    ReadBHToRun(out string[] bhNames, out double[] reduceLevels, out int[] colNums);
                    #endregion

                    #region Read Excel Info
                    Range soilInputData = ((RangeTextBox)TextBoxAttributeDic["inputSoil_pile"]).GetRangeFromFullAddress();
                    Worksheet refSheet = ((SheetTextBox)TextBoxAttributeDic["refSheet_pile"]).getSheet();
                    // Rock type to hashset
                    HashSet<string> rockTypes = new HashSet<string>();
                    Range rockTypeRange = ((RangeTextBox)TextBoxAttributeDic["inputRockTypes_pile"]).GetRangeFromFullAddress();
                    foreach(Range cell in rockTypeRange)
                    {
                        if (cell.Value2 != null)
                        {
                            rockTypes.Add(cell.Value2.ToString());
                        }
                    }

                    // NSF type to hashset
                    HashSet<string> nsfType = new HashSet<string>();
                    if (((RangeTextBox)TextBoxAttributeDic["inputNsfTypes_pile"]).textBox.Text != "")
                    {
                        Range nsfTypeRange = ((RangeTextBox)TextBoxAttributeDic["inputNsfTypes_pile"]).GetRangeFromFullAddress();
                        foreach (Range cell in nsfTypeRange)
                        {
                            if (cell.Value2 != null)
                            {
                                nsfType.Add(cell.Value2.ToString());
                            }
                        }
                    }
                    #endregion

                    #region Loop through sheets and Copy Values
                    for (int bhIndex = 0; bhIndex < bhNames.Length; bhIndex++)
                    {
                        string bhName = bhNames[bhIndex];
                        progressTracker.UpdateStatus($"Copying {bhName}{dispAppendName.Text}");
                        worker.ReportProgress(ConvertToProgress(bhIndex, bhNames.Length));

                        CopyDataToNewSheet(bhIndex);

                        if (worker.CancellationPending)
                        {
                            break;
                        }
                    }
                    #endregion
                    #endregion

                    #region Helper Functions
                    void ReadBHToRun(out string[] bhNames_L, out double[] reduceLevels_L, out int[] colNums_L)
                    {
                        Range soilInput = ((RangeTextBox)TextBoxAttributeDic["inputSoil_pile"]).GetRangeFromFullAddress();
                        bhNames_L = new string[soilInput.Columns.Count / 3];
                        reduceLevels_L = new double[soilInput.Columns.Count / 3];
                        colNums_L = new int[soilInput.Columns.Count / 3];

                        for (int i = 0; i < bhNames_L.Length; i++)
                        {
                            colNums_L[i] = (i) * 3 + 1;
                            if (soilInput.Cells[1, colNums_L[i]].Value2 == null)
                            {
                                continue;
                            }
                            bhNames_L[i] = soilInput.Cells[1, colNums_L[i]].Value2.ToString();
                            reduceLevels_L[i] = ReadDoubleFromCell(soilInput.Cells[2, colNums_L[i] + 1]) + 100;
                        }

                        #region Delete Sheets if Existing
                        // Check if sheet already exist
                        List<string> existingWorksheets = new List<string>();
                        string existingWorksheetString = "";
                        foreach (string bhName_L in bhNames_L)
                        {
                            try
                            {
                                string sheetName = bhName_L + dispAppendName.Text;
                                Worksheet worksheet = ThisWorkBook.Sheets[sheetName];
                                existingWorksheets.Add(sheetName);
                                existingWorksheetString += sheetName + "\n";
                            }
                            catch //(Exception ex)
                            {
                                // Do nothing
                            }
                        }

                        if (existingWorksheets.Count > 0)
                        {
                            DialogResult result = MessageBox.Show($"The following sheets already exist, delete these sheets?\n{existingWorksheetString}", "Confirmation", MessageBoxButtons.YesNoCancel);
                            if (result == DialogResult.Yes)
                            {
                                try
                                {
                                    ThisApplication.DisplayAlerts = false;
                                    foreach (string sheetName in existingWorksheets)
                                    {
                                        Worksheet worksheet = ThisWorkBook.Sheets[sheetName];
                                        worksheet.Delete();
                                    }
                                }
                                finally
                                {
                                    ThisApplication.DisplayAlerts = true;
                                }
                            }
                            else if (result == DialogResult.Cancel)
                            {
                                throw new Exception("Terminated by user");
                            }
                        }
                        #endregion
                    }

                    void CopyDataToNewSheet(int bhIndex)
                    {
                        #region Copy Sheet
                        string bhName = bhNames[bhIndex];
                        int colNum = colNums[bhIndex];
                        double reduceLevel = reduceLevels[bhIndex];
                        Worksheet newSheet = null;
                        string sheetName = bhName + dispAppendName.Text;

                        try 
                        {
                            // if sheet exists
                            newSheet = ThisWorkBook.Sheets[sheetName];
                            Range soilRange = ((RangeTextBox)TextBoxAttributeDic["soilDest_pile"]).GetRangeForSpecificSheet(sheetName);
                            Range editRange = soilRange.Resize[soilRange.Rows.Count - 1, soilRange.Columns.Count].Offset[1, 0];
                            for (int i = 1; (i <= editRange.Columns.Count && i <= 4); i++)
                            {
                                Range cellsToClear = editRange.Columns[i];
                                cellsToClear.ClearContents();
                            }
                        }
                        catch
                        {
                            // if sheet doesnt exist
                            newSheet = CopyNewSheetAtBack(refSheet, sheetName);
                        }
                        newSheet.Name = sheetName;
                        ((RangeTextBox)TextBoxAttributeDic["bhRLCell_pile"]).GetRangeForSpecificSheet(sheetName).Value2 = reduceLevel;
                        #endregion

                        #region Read Input Range
                        Range startCell = soilInputData.Cells[6, colNum];
                        Range endCell = soilInputData.Cells[soilInputData.Rows.Count, colNum + 2];
                        Range thisSoilInputData = soilInputData.Worksheet.Range[startCell, endCell];
                        Range thisRockStartInput = soilInputData.Cells[4, colNum + 2];
                        Range thisSpt100StartInput = soilInputData.Cells[4, colNum + 1];
                        #endregion

                        #region Copy Soil Data
                        Range soilDestination = ((RangeTextBox)TextBoxAttributeDic["soilDest_pile"]).GetRangeForSpecificSheet(sheetName);
                        Range thisDepth = soilDestination.Columns[1];
                        Range thisSoilType = soilDestination.Columns[2];
                        Range thisSptN = soilDestination.Columns[3];
                        Range thisVeff = soilDestination.Columns[5];
                        Range thieBeta = soilDestination.Columns[6];
                        Range thisNsf= soilDestination.Columns[9];
                        Range thisRockStartDest = ((RangeTextBox)TextBoxAttributeDic["rockStart_pile"]).GetRangeForSpecificSheet(sheetName);
                        Range thisSpt100StartDest = ((RangeTextBox)TextBoxAttributeDic["spt100Start_pile"]).GetRangeForSpecificSheet(sheetName);
                        Range nsfFormulaCell = soilDestination.Cells[3, 5];
                        Range betaFormulaCell = soilDestination.Cells[3, 6];

                        double firstDepth = ReadDoubleFromCell(soilDestination.Cells[1, 1]);
                        int outSoilIndex = 3; //start from row 3
                        
                        for (int inSoilIndex = 1; inSoilIndex <= soilInputData.Rows.Count; inSoilIndex++)
                        {
                            #region Checks
                            //Break if it is empty row
                            if (thisSoilInputData.Cells[inSoilIndex, 1].Value2 == null)
                            {
                                break;
                            }
                            else if (thisSoilInputData.Cells[inSoilIndex, 1].Value2.ToString() == "")
                            {
                                break;
                            }
                            // Skip if depth is above the first depth after COL
                            if (ReadDoubleFromCell(thisSoilInputData.Cells[inSoilIndex, 1]) < firstDepth)
                            {
                                continue;
                            }
                            // Skip if SPT == "";
                            if (thisSoilInputData.Cells[inSoilIndex + 1, 1].Value2 != null) // Don't skip last row
                                {
                                if (thisSoilInputData.Cells[inSoilIndex, 2].Value2 == null || thisSoilInputData.Cells[inSoilIndex, 2].Text == "")
                                {
                                    continue;
                                }
                            }
                            // Throw error if last line is created
                            if (outSoilIndex >= soilDestination.Rows.Count + 1)
                            {
                                Beaver.LogError($"Input soil data exceeds size of destination for {bhName}. Unable to copy all information");
                            }
                            #endregion

                            #region Copy Values
                            //Copy Depth to previous layer (start depth from input is "end depth")
                            thisDepth.Cells[outSoilIndex-1].Value2 = thisSoilInputData.Cells[inSoilIndex, 1].Value2;

                            #region Soil Type and NSF
                            //Copy Soil Type 
                            thisSoilType.Cells[outSoilIndex].Value2 = thisSoilInputData.Cells[inSoilIndex, 3].Value2;
                            string currentSoilType = thisSoilType.Cells[outSoilIndex].Value2;

                            // Set NSF
                            if (nsfType.Contains(currentSoilType)) // Is nsfType
                            {
                                thisNsf.Cells[outSoilIndex].Value2 = "y";
                                // Set formula
                                if (outSoilIndex > 3)
                                {
                                    nsfFormulaCell.Copy(thisVeff.Cells[outSoilIndex]);
                                    betaFormulaCell.Copy(thieBeta.Cells[outSoilIndex]);
                                }
                            }
                            else
                            {
                                nsfFormulaCell.ClearComments();
                                thisNsf.Cells[outSoilIndex].Value2 = "n";
                            }
                            #endregion

                            #region SPT 
                            //Copy SPT
                            if (rockTypes.Contains(currentSoilType)) // Is rock
                            {
                                thisSptN.Cells[outSoilIndex].Value2 = "R";
                            }
                            else
                            {
                                thisSptN.Cells[outSoilIndex].Value2 = thisSoilInputData.Cells[inSoilIndex, 2].Value2;
                            }

                            //If we are at cell 2, copy SPT number 
                            if (outSoilIndex == 3)
                            {
                                // Find previous SPT value if any
                                double previousSptValue = 0;
                                for (int previousSptIndex = inSoilIndex-1; previousSptIndex >=1; previousSptIndex--)
                                {
                                    if (thisSoilInputData.Cells[previousSptIndex, 2].Value2 != null)
                                    {
                                        previousSptValue = thisSoilInputData.Cells[previousSptIndex, 2].Value2;
                                        break;
                                    }
                                }
                                thisSptN.Cells[outSoilIndex - 1].Value2 = previousSptValue;

                                // Find previous soil type, if any
                                string previousSoilType = "";
                                for (int previousSptIndex = inSoilIndex - 1; previousSptIndex >= 1; previousSptIndex--)
                                {
                                    if (thisSoilInputData.Cells[previousSptIndex, 3].Value2 != null)
                                    {
                                        previousSoilType = thisSoilInputData.Cells[previousSptIndex, 3].Value2;
                                        break;
                                    }
                                }
                                thisSoilType.Cells[outSoilIndex - 1].Value2 = previousSoilType;
                            }
                            #endregion

                            outSoilIndex++;
                            #endregion
                        }
                        #endregion

                        #region Copy Rock & SPT100 Start
                        try
                        {
                            if (thisRockStartInput.Value2 != null)
                            {
                                thisRockStartDest.Value2 = thisRockStartInput.Value2;
                            }
                            else
                            {
                                thisRockStartDest.Value2 = "NA";
                            }
                        }
                        catch (Exception ex)
                        {
                            Beaver.LogError($"Unable to read rock start value at {thisRockStartInput.Worksheet.Name}!{thisRockStartInput.Address} for {bhName}\n    " + ex.Message);
                        }

                        try
                        {
                            if (thisSpt100StartInput.Value2 != null)
                            {
                                thisSpt100StartDest.Value2 = thisSpt100StartInput.Value2;
                            }
                            else
                            {
                                thisSpt100StartDest.Value2 = "NA";
                            }
                        }
                        catch (Exception ex)
                        {
                            Beaver.LogError($"Unable to read SPT100 start value at {thisSpt100StartInput.Worksheet.Name}!{thisSpt100StartInput.Address} for {bhName}\n    " + ex.Message);
                        }
                        #endregion

                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }
                finally
                {
                    ThisApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                    ThisApplication.ScreenUpdating = true;
                }
                
                // for each (xxx)
                // read bh info, return obj representing BH
                // pass this obj to another function to create sheet and copy obj. return sheet to run
                // pass sheet to optimisation function
                // finish                
            });
        }
        #endregion

        #region Design Piles
        private void designPiles_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                Beaver.InitializeForWorkbook(ThisApplication.ActiveWorkbook);
                try
                {
                    if (checkDeactivateScreen.Checked)
                    {
                        ThisApplication.ScreenUpdating = false;
                    }
                    // Get Sheet Names
                    string[] sheetsNames = GetSheetsToRun();

                    // Design each sheet
                    int prog = 0;
                    int maxprog = sheetsNames.Length;
                    foreach (string sheetName in sheetsNames)
                    {
                        worker.ReportProgress(ConvertToProgress(prog, maxprog));
                        DesignSheet(sheetName);
                        prog++;
                        if (worker.CancellationPending)
                        {
                            break;
                        }
                    }

                    if (worker.CancellationPending)
                    {
                        return;
                    }
                    progressTracker.UpdateStatus("Completed");
                    MessageBox.Show("Completed", "Completed");

                    #region Sub Functions
                    string[] GetSheetsToRun()
                    {
                        ////Checks that all sheets are valid
                        //Range selectedSheetRange = ((RangeTextBox)RangeAttributeDic["sheetsToRun_pile"]).GetRangeFromFullAddress();

                        //List<string> sheetsToDesign = new List<string>();
                        //foreach (Range cell in selectedSheetRange)
                        //{
                        //    if (cell.Value2 == null)
                        //    {
                        //        continue;
                        //    }
                        //    string sheetName = cell.Value2.ToString();
                        //    try
                        //    {
                        //        Worksheet thisSheet = ThisWorkBook.Worksheets[sheetName];
                        //        sheetsToDesign.Add(sheetName);
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        throw new Exception($"Unable to find worksheet '{sheetName}'\n\n" + ex.Message);
                        //    }
                        //}

                        HashSet<string> sheetsToDesign = ((MultipleSheetsAttribute)OtherAttributeDic["sheetsToRun_pile"]).GetSheetNamesHash();
                        //HashSet<string> sheetsToDesign = new HashSet<string>();
                        //sheetsToDesign.Add("BH 01");
                        return sheetsToDesign.ToArray();
                    }
                    
                    void DesignSheet(string sheetName)
                    {
                        #region Set Excel Range
                        Range soilDataRange = ((RangeTextBox)TextBoxAttributeDic["soilDest_pile"]).GetRangeForSpecificSheet(sheetName);
                        Range fsRange = ((RangeTextBox)TextBoxAttributeDic["fsRange_pile"]).GetRangeForSpecificSheet(sheetName);
                        //Range fsWorkingRange = fsRange.Resize[fsRange.Rows.Count - 2, fsRange.Columns.Count].Offset[2, 0];
                        Range qbRange = ((RangeTextBox)TextBoxAttributeDic["qbRange_pile"]).GetRangeForSpecificSheet(sheetName);
                        Range rockStartCell = ((RangeTextBox)TextBoxAttributeDic["rockStart_pile"]).GetRangeForSpecificSheet(sheetName);
                        Range effRange = ((RangeTextBox)TextBoxAttributeDic["effRange_pile"]).GetRangeForSpecificSheet(sheetName);
                        //Range referenceFormula = fsWorkingRange.Cells[0, 1];
                        Range referenceFormula = fsRange.Cells[2, 1];
                        #endregion
                        #region Clear fsRange (row 3 onwards)
                        Range clearRange = fsRange.Resize[fsRange.Rows.Count - 2, fsRange.Columns.Count].Offset[2, 0];
                        clearRange.ClearContents();
                        clearRange = null;
                        //fsWorkingRange.ClearContents();
                        #endregion


                        #region Loop through fsRange
                        int workingRownum = -1;
                        for (int colNum = 1; colNum <= fsRange.Columns.Count; colNum++) // Loops through pile sizes
                        {
                            if (worker.CancellationPending)
                            {
                                break;
                            }

                            #region Initialise
                            Range efficiencyCell = ((RangeTextBox)TextBoxAttributeDic["effRange_pile"]).GetRangeForSpecificSheet(sheetName).Cells[colNum, 6];
                            double efficiencyLowerBound = TextBoxAttributeDic["effLower_pile"].GetDoubleFromTextBox();
                            double efficiencyUpperBound = TextBoxAttributeDic["effUpper_pile"].GetDoubleFromTextBox();
                            bool efficiencyMet = false;
                            if (!checkDeactivateScreen.Checked)
                            {
                                fsRange.Worksheet.Activate();
                                fsRange.Cells[1,1].Select();
                            }
                            #endregion

                            #region Find first passing value
                            for (int rowNum = 3; rowNum <= fsRange.Rows.Count; rowNum++) // Loops through each soil layer in fs
                            {
                                progressTracker.UpdateStatus($"Designing sheet: {sheetName}, pile size {effRange.Cells[colNum, 1].Value2.ToString()}");
                                #region Copy until new layer
                                if (colNum > 1 && rowNum <= workingRownum)
                                {
                                    try
                                    {
                                        ThisApplication.Calculation = XlCalculation.xlCalculationManual;
                                        Range startCell = fsRange.Cells[3,colNum];
                                        Range endCell = fsRange.Cells[workingRownum, colNum];

                                        Range targetRange = fsRange.Worksheet.Range[startCell, endCell];
                                        targetRange.FormulaR1C1 = referenceFormula.FormulaR1C1;
                                        rowNum = workingRownum + 1;
                                    }
                                    finally
                                    {
                                        ThisApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                                    }
                                }
                                #endregion

                                Range editCell = fsRange.Cells[rowNum, colNum];
                                editCell.FormulaR1C1 = referenceFormula.FormulaR1C1;
                                //referenceFormula.Copy(editCell);

                                #region Check if layer is rock
                                CheckRock(colNum, rowNum);
                                #endregion

                                #region Check if efficiency is met
                                try
                                {
                                    double efficiencyValue = ReadDoubleFromCell(efficiencyCell);
                                    if (efficiencyValue < efficiencyUpperBound)
                                    {
                                        efficiencyMet = true;
                                        //workingRownum = rowNum + 2; // set row where we want to further optimise, +2 to be relative to fs Range
                                        workingRownum = rowNum;
                                        break;
                                    }
                                }
                                catch //(Exception ex)
                                {
                                    efficiencyMet = false;
                                }
                                #endregion
                            }

                            #region Log Error
                            if (efficiencyMet == false)
                            {
                                // Log cases that are not optimised
                                Beaver.LogError($"Error: Unable to find optimum value for {sheetName}, column {colNum}, pile size {effRange.Cells[colNum,1].Value2.ToString()} and above (if any).\n" +
                                    $"Maximum number of rows reached for fs");
                                break; // Break loop for this sheet, continue to next sheet
                            }
                            #endregion
                            #endregion

                            #region Create new row in soil data

                            if (ReadDoubleFromCell(efficiencyCell) >= efficiencyLowerBound)
                            {
                                continue; // No further optimisation required
                            }

                            #region Find last valid cell in Soil Data
                            int maxRowNum = soilDataRange.Rows.Count;
                            for (int rowNum = soilDataRange.Rows.Count; rowNum >= 0; rowNum--)
                            {
                                Range cell = soilDataRange.Cells[rowNum, 1];
                                if ((cell.Value2 != null) && (cell.Text != ""))
                                {
                                    maxRowNum = rowNum;
                                    break;
                                }
                            }

                            if (maxRowNum == soilDataRange.Rows.Count) // Check if soil data is full
                            {
                                Beaver.LogError($"Error: Unable to find optimum value for {sheetName}, column {colNum}, pile size {effRange.Cells[colNum, 1].Value2.ToString()} and above (if any).\n" +
                                    $"Maximum number of rows reached for soil data");
                                break; // Break loop for this sheet, continue to next sheet
                            }
                            #endregion

                            #region Copy cells 

                            // From last valid cell to working row number, copy value/formula down by 1
                            try
                            {
                                ThisApplication.Calculation = XlCalculation.xlCalculationManual;
                                Range startCell = soilDataRange.Cells[workingRownum, 1];
                                Range endCell = soilDataRange.Cells[maxRowNum, 3];
                                Range moveSource = soilDataRange.Worksheet.Range[startCell, endCell];
                                Range moveDest = moveSource.Offset[1, 0];
                                moveDest.Value2 = moveSource.Value2;

                                startCell = soilDataRange.Cells[workingRownum, 4];
                                endCell = soilDataRange.Cells[maxRowNum, 9];
                                moveSource = soilDataRange.Worksheet.Range[startCell, endCell];
                                moveDest = moveSource.Offset[1, 0];
                                moveDest.FormulaR1C1 = moveSource.FormulaR1C1;

                                
                                //for (int rowNum = maxRowNum; rowNum >= workingRownum; rowNum--)
                                //{
                                //    Range row = soilDataRange.Rows[rowNum];
                                //    for (int colNum_local = 1; colNum_local <= row.Columns.Count; colNum_local++)
                                //    {
                                //        Range cell = row.Cells[1, colNum_local];
                                //        Range destCell = cell.Offset[1, 0];
                                //        if (colNum_local <= 4)
                                //        {
                                //            destCell.Value2 = cell.Value2;
                                //        }
                                //        else
                                //        {
                                //            cell.Copy(destCell);
                                //        }
                                //    }
                                //}
                            }
                            finally
                            {
                                ThisApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                            }

                            // For depth value in first cell, copy from cell above
                            Range currentDepthCell = soilDataRange.Cells[workingRownum,1];
                            currentDepthCell.Value2 = currentDepthCell.Offset[-1, 0].Value2;
                            ThisApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                            #endregion

                            #endregion

                            #region Optimise new row
                            double effValue;
                            try
                            {
                                effValue = ReadDoubleFromCell(efficiencyCell);
                            }
                            catch
                            {
                                effValue = double.PositiveInfinity;
                            }

                            while (ReadDoubleFromCell(currentDepthCell)< ReadDoubleFromCell(currentDepthCell.Offset[1,0])) 
                            {
                                currentDepthCell.Value2 += 0.1;
                                CheckRock(colNum, workingRownum);

                                try
                                {
                                    effValue = ReadDoubleFromCell(efficiencyCell);
                                }
                                catch
                                {
                                    effValue = double.PositiveInfinity;
                                }

                                if (effValue < efficiencyUpperBound)
                                {
                                    break;
                                }
                            }
                            #endregion

                            if (worker.CancellationPending)
                            {
                                break;
                            }
                            #region SubFunction
                            void CheckRock(int colNum_Local, int rowNum)
                            {
                                //Range currentDepthCell = effRange.Cells[colNum_Local, 7];
                                Range failRock = effRange.Cells[colNum_Local, 9]; // Rows and columns are swapped for this table
                                try
                                {
                                    if (!failRock.Value2) // sufficient rock socketing
                                    {
                                        qbRange.Cells[1, colNum_Local].Value2 = "R";
                                    }
                                    else
                                    {
                                        if (soilDataRange[rowNum, 3].Value2.ToString() == "R")
                                        {
                                            qbRange.Cells[1, colNum_Local].Value2 = 100;
                                        }
                                        else
                                        {
                                            qbRange.Cells[1, colNum_Local].Value2 = soilDataRange[rowNum, 3].Value2;
                                        }
                                    }
                                }
                                catch
                                {
                                    if (soilDataRange[rowNum, 3].Value2 == "R")
                                    {
                                        qbRange.Cells[1, colNum_Local].Value2 = 100;
                                    }
                                    else
                                    {
                                        qbRange.Cells[1, colNum_Local].Value2 = soilDataRange[rowNum, 3].Value2;
                                    }
                                }
                            }
                            #endregion
                        }
                        #endregion
                    }
                    

                    #endregion
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    ThisApplication.Calculation = XlCalculation.xlCalculationAutomatic;
                    ThisApplication.ScreenUpdating = true ;
                    Beaver.CheckLog();
                }
            });
                
        }

        #endregion

        //private void setSheetsToRun_Click(object sender, EventArgs e)
        //{
        //    string AttName = "sheetsToRun_pile";
        //    try
        //    {
        //        using (SheetSelector sheetSelector = new SheetSelector(AttName))
        //        {
        //            sheetSelector.ShowDialog();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error");
        //    }
        //}
    }
}

