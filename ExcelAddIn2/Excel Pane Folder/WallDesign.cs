using ETABSv1;
using Microsoft.Office.Interop.Excel;
using MigraDoc.DocumentObjectModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Color = System.Drawing.Color;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class WallDesign : UserControl
    {
        #region Init
        Dictionary<string, AttributeTextBox> TextBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();
        public WallDesign()
        {
            InitializeComponent();
            CreateAttributes();
            AddToolTips();
        }

        private void AddToolTips()
        {
            #region MyRegion

            #endregion

            #region Design Rebar
            toolTip1.SetToolTip(overwriteRebarCheck,
                "If unchecked, initial check will be done based on current values in the output range\n" +
                "If checked, initial check will be done based on values matched from Rebar Table");
            #endregion

            #region Additional Settings
            toolTip1.SetToolTip(backupSheetCheck,
                "If checked, new sheet with will be copied and added at the back\n" +
                "This will delete and overwite any existing sheet with the same sheetname\n" +
                "Sheet name used = Current sheet name + \"_backup\"\n" +
                "Sheets duplicated:\n" +
                "  Pier Label Range worksheet (match, design, unify changes)\n" +
                "  Rebar Table worksheet (unify changes)");
            #endregion
        }

        private void CreateAttributes()
        {
            #region Match Reinforcement
            AttributeTextBox attTB = new RangeTextBox("rebarTable_WD", dispRebarTable, setRebarTable, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("storeyTable_WD", dispStoreyTable, setStoreyTable, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("pierLabelRange_WD", dispPierLabelRange, setPierLabelRange, "column");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("matchStoreyRange_WD", dispMatchStoreyCol, setMatchStoreyCol, "cell");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("outputRange_WD", dispOutputCol, setOutputCol, "cell");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("statusRange_WD", dispStatusCol, setStatusCol, "cell");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            var att = new CheckBoxAttribute("overwriteInitialRebar_WD", overwriteRebarCheck);
            att = new CheckBoxAttribute("overwriteInitialRebar_WD", unifyChangesCheck);
            #endregion

            #region Modify Reinforcement
            attTB = new RangeTextBox("mainRebarHeirarchy_WD", dispRebarHeirarchy, setRebarHeirarchy, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("targetUR_WD", dispTargetUR, true);
            attTB.type = "double";
            attTB.SetDefaultValue("0.9");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("maxAs_WD", dispMaxAs, true);
            attTB.type = "double";
            attTB.SetDefaultValue("4");
            TextBoxAttributeDic.Add(attTB.attName, attTB);


            #endregion

            #region Additional Settings
            att = new CheckBoxAttribute("backupSheetCheck_WD", backupSheetCheck);
            att = new CheckBoxAttribute("resetFontColorRebarTable_WD", resetFontColourRebarTableCheck);
            att = new CheckBoxAttribute("resetFontColorCheckSheet_WD", resetFontColourCheckSheetCheck);
            #endregion

            #region Decomposer
            attTB = new RangeTextBox("decomposeRange_WD", dispDecomposeRange, setDecomposeRange, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);
            #endregion
        }

        #endregion

        #region Match Rebars
        private void matchWallRebar_Click(object sender, EventArgs e)
        {
            List<TrackedRange> trackedRanges = null;
            try
            {
                Stopwatch totalStopwatch = Stopwatch.StartNew();
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                trackedRanges = CreateMatchTrackRanges();
                CreatePierLabelBackupSheet();
                ReadStoreyTable();
                ReadRebarTable();

                MatchRebars();

                Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                pierLabelRange.Worksheet.Activate();

                totalStopwatch.Stop();
                MessageBox.Show($"Total Execution Time: {totalStopwatch.ElapsedMilliseconds} ms", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                HighlightChangesForMatchRebar(trackedRanges);
                #region Release Dictionaries
                rebarDic = null;
                storeyTracker = null;
                #endregion
            }
        }

        private void MatchRebars()
        {
            try
            {
                #region Get Ranges
                Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                Range storeyRange = ((RangeTextBox)TextBoxAttributeDic["matchStoreyRange_WD"]).GetRangeFromFullAddress();

                Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
                outputRange = outputRange.Worksheet.Cells[pierLabelRange.Row, outputRange.Column];
                Range statusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();
                statusRange = statusRange.Worksheet.Cells[pierLabelRange.Row, statusRange.Column];

                #region Read Ranges to arrays
                string[] pierLabels = GetContentsAsStringArray(pierLabelRange, false);
                int pierLabelStartRow = pierLabelRange.Row;
                string[] etabsStoreyNames;
                {
                    Range etabsStoreyRange = GetColRangeFromRanges(pierLabelRange, storeyRange);
                    etabsStoreyNames = GetContentsAsStringArray(etabsStoreyRange, false);
                }
                #endregion

                #endregion

                #region Init Arrays
                int outputLength = pierLabelRange.Rows.Count;
                double[] rebarDia = new double[outputLength];
                double[] rebarSpacing = new double[outputLength];
                double[] shearDia = new double[outputLength];
                double[] shearSpacing = new double[outputLength];
                string[] status = new string[outputLength];
                #endregion

                #region Match
                for (int rowNum = 0; rowNum < pierLabels.Length; rowNum++)
                {
                    string pierLabel = pierLabels[rowNum];
                    string etabsStoreyName = etabsStoreyNames[rowNum];
                    MatchSingleRow(rowNum, pierLabels[rowNum], etabsStoreyNames[rowNum],
                ref status[rowNum], ref rebarDia[rowNum], ref rebarSpacing[rowNum], ref shearDia[rowNum], ref shearSpacing[rowNum]);
                }
                #endregion

                #region Write To Excel
                WriteToExcelRangeAsCol(outputRange, 0, 0, false, rebarDia, rebarSpacing);
                WriteToExcelRangeAsCol(outputRange, 0, 8, false, shearDia, shearSpacing);
                WriteToExcelRangeAsCol(statusRange, 0, 0, false, status);
                #endregion
            }
            catch (Exception ex) { throw new Exception("Error matching reinforcement\n" + ex.Message); }
            
        }

        private void MatchSingleRow(int rowNum, string pierLabel, string etabsStoreyName,
            ref string status, ref double rebarDia, ref double rebarSpacing, ref double shearDia, ref double shearSpacing)
        {
            try
            {
                if (!rebarDic.ContainsKey(pierLabel))
                {
                    status = $"Error: Wall Label {pierLabel} not found in rebar table " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
                    return;
                }

                AssignedWallRebar wallRebar = rebarDic[pierLabel];
                object[] rowValues = wallRebar.GetStoreyData(storeyTracker.GetStoreyNum(etabsStoreyName, "etabs"));

                rebarDia = (double)rowValues[3];
                rebarSpacing = (double)rowValues[4];
                shearDia = (double)rowValues[5];
                shearSpacing = (double)rowValues[6];
                status = "Matched rebars " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            }
            catch (Exception ex)
            {
                status = "Error: " + ex.Message + " " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString(); ;
            }
        }

        private List<TrackedRange> CreateMatchTrackRanges()
        {
            List<TrackedRange> trackedRanges = new List<TrackedRange>();
            
            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            
            Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
            outputRange = outputRange.Worksheet.Cells[pierLabelRange.Row, outputRange.Column];

            Range rebarDiaRange = GetColRangeFromRanges(pierLabelRange, outputRange);
            Range rebarSpacingRange = GetColRangeFromRanges(pierLabelRange, outputRange, 0 ,1);
            Range shearDiaRange = GetColRangeFromRanges(pierLabelRange, outputRange, 0, 8);
            Range shearSpacingRange = GetColRangeFromRanges(pierLabelRange, outputRange, 0, 9);

            TrackedRange trackedRange = new TrackedRange(rebarDiaRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);
            trackedRange = new TrackedRange(rebarSpacingRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);
            trackedRange = new TrackedRange(shearDiaRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);
            trackedRange = new TrackedRange(shearSpacingRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);

            return trackedRanges;
        }

        private void HighlightChangesForMatchRebar(List<TrackedRange> trackedRanges)
        {
            HighlightChangesForTrackedRanges(trackedRanges);

            Range statusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();
            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            Range matchStatusRange = GetColRangeFromRanges(pierLabelRange, statusRange, 0, 0);
            HighlightErrors(matchStatusRange, true);
        }
        #endregion

        #region Shared: Reading Reference Tables
        Dictionary<string, AssignedWallRebar> rebarDic;
        private void ReadRebarTable()
        {
            try
            {
                Range rebarTableRange = ((RangeTextBox)TextBoxAttributeDic["rebarTable_WD"]).GetRangeFromFullAddress();
                rebarDic = new Dictionary<string, AssignedWallRebar>();
                foreach (Range row in rebarTableRange.Rows)
                {
                    string name = row.Cells[1].Text;
                    if (!rebarDic.ContainsKey(name)) { rebarDic[name] = new AssignedWallRebar(name, storeyTracker); }
                    AssignedWallRebar wallRebar = rebarDic[name];
                    wallRebar.AddRow(row);
                }

                foreach (AssignedWallRebar wallRebar in rebarDic.Values)
                {
                    wallRebar.SortStories();
                }
            }
            catch (Exception ex) { throw new Exception("Error reading rebar table\n" + ex.Message); }
        }

        StoreyTracker storeyTracker;
        private void ReadStoreyTable()
        {
            try
            {
                Range storeyTable = ((RangeTextBox)TextBoxAttributeDic["storeyTable_WD"]).GetRangeFromFullAddress();
                storeyTracker = new StoreyTracker(storeyTable);
            }
            catch (Exception ex) { throw new Exception("Error reading storey table\n" + ex.Message); }
        }

        RebarHeirarchy rebarHeirarchy;
        private void ReadRebarHeirarchy()
        {
            try
            {
                Range rebarHeirarchyRange = ((RangeTextBox)TextBoxAttributeDic["mainRebarHeirarchy_WD"]).GetRangeFromFullAddress();
                rebarHeirarchy = new RebarHeirarchy(rebarHeirarchyRange);
            }
            catch (Exception ex) { throw new Exception("Error reading rebar heirarchy table\n" + ex.Message); }
        }

        private void CreateRebarTableBackupSheet()
        {
            if (!backupSheetCheck.Checked) { return; }
            Range rebarTableRange = ((RangeTextBox)TextBoxAttributeDic["rebarTable_WD"]).GetRangeFromFullAddress();

            Worksheet worksheet = rebarTableRange.Worksheet;
            CopyNewSheetAtBack(worksheet, worksheet.Name + "_backup", true);
        }
        private void CreatePierLabelBackupSheet()
        {
            if (!backupSheetCheck.Checked) { return; }
            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();

            Worksheet worksheet = pierLabelRange.Worksheet;
            CopyNewSheetAtBack(worksheet, worksheet.Name + "_backup", true);
        }
        #endregion

        #region Shared: Highlight Changes
        private void HighlightChangesForTrackedRanges(List<TrackedRange> trackedRanges)
        {
            foreach (TrackedRange trackedRange in trackedRanges)
            {
                trackedRange.HighlightChanges();
            }
        }
        private void HighlightErrors(Range range, bool resetRange)
        {
            if (resetRange)
            {
                range.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            }
            List<Range> formatRanges = new List<Range>();
            foreach (Range cell in range)
            {
                if (cell.Value2 == null) { continue; }
                string text = cell.Value2.ToString();
                if (text.Length < 5) { continue; }
                if (text.Substring(0, 5) == "Error")
                {
                    cell.Font.Color = Color.Red;
                }
            }
        }
        #endregion

        #region Design Rebar
        private void designRebar_Click(object sender, EventArgs e)
        {
            List<TrackedRange> trackedRanges = new List<TrackedRange>();
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Stopwatch totalStopwatch = Stopwatch.StartNew();

                CreatePierLabelBackupSheet();
                if (unifyChangesCheck.Checked){
                    List<TrackedRange> unifyTrackRange = CreateUnifyTrackRanges();
                    trackedRanges.AddRange(unifyTrackRange);
                    CreateRebarTableBackupSheet(); 
                }

                ReadStoreyTable();
                ReadRebarTable();
                ReadRebarHeirarchy();

                if (overwriteRebarCheck.Checked){ MatchRebars(); }
                trackedRanges.AddRange(CreateDesignTrackRanges());
                
                SolveForRebar();
                if (unifyChangesCheck.Checked) { UnifyChanges(); }
                
                Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                pierLabelRange.Worksheet.Activate();

                totalStopwatch.Stop();
                MessageBox.Show($"Total Execution Time: {totalStopwatch.ElapsedMilliseconds} ms", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                HighlightDesignChanges(trackedRanges);
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                rebarHeirarchy = null;
                rebarDic = null;
                storeyTracker = null;
            }
        }

        private void SolveForRebar()
        {
            try
            {
                #region Get Ranges
                Range rebarHeirarchyRange = ((RangeTextBox)TextBoxAttributeDic["mainRebarHeirarchy_WD"]).GetRangeFromFullAddress();

                Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                Range storeyRange = ((RangeTextBox)TextBoxAttributeDic["matchStoreyRange_WD"]).GetRangeFromFullAddress();
                Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
                outputRange = outputRange.Worksheet.Cells[pierLabelRange.Row, outputRange.Column];

                Range checkSheetStatusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();
                checkSheetStatusRange = checkSheetStatusRange.Worksheet.Cells[pierLabelRange.Row, checkSheetStatusRange.Column];

                double targetUr = TextBoxAttributeDic["targetUR_WD"].GetDoubleFromTextBox();
                if (targetUr > 1) { throw new ArgumentException("Target UR cannot be greater than 1"); }

                double maxAsPer = TextBoxAttributeDic["maxAs_WD"].GetDoubleFromTextBox();
                if (maxAsPer <= 0) { throw new ArgumentException("Max As cannot be smaller than or equal to 0"); }
                #endregion

                #region Create Status Range
                string[] rebarStatus = new string[rebarHeirarchyRange.Rows.Count];
                string[] checkSheetStatus = new string[pierLabelRange.Rows.Count];
                #endregion

                #region Iterate through each row
                Worksheet outputSheet = outputRange.Worksheet;
                string[] optimisationStatus = new string[pierLabelRange.Rows.Count];
                for (int rowNum = 0; rowNum < pierLabelRange.Rows.Count; rowNum++)
                {
                    #region Check Bending
                    DesignBendingRebarForRow();

                    void DesignBendingRebarForRow()
                    {
                        try
                        {
                            Range asPercRange = outputRange.Offset[rowNum, 4];
                            Range urRange = outputRange.Offset[rowNum, 5];
                            double asPerc = double.Parse(asPercRange.Value2.ToString());
                            double ur = double.Parse(urRange.Value2.ToString());
                            optimisationStatus[rowNum] = "";

                            if (ur <= targetUr && asPerc < maxAsPer) { optimisationStatus[rowNum] = "Ok"; return; }

                            Range diaRange = outputRange.Offset[rowNum, 0];
                            Range spacingRange = outputRange.Offset[rowNum, 1];

                            int counter = 0;
                            bool rebarDecreased = false;
                            while (counter < rebarHeirarchy.table.GetLength(0))
                            {
                                Globals.ThisAddIn.Application.Calculate();
                                asPerc = double.Parse(asPercRange.Value2.ToString());
                                ur = double.Parse(urRange.Value2.ToString());

                                double dia = (double)outputRange.Offset[rowNum, 0].Value2;
                                double spacing = (double)outputRange.Offset[rowNum, 1].Value2;

                                if (asPerc >= maxAsPer)
                                {
                                    if (ur > targetUr)
                                    {
                                        optimisationStatus[rowNum] = "Error: No solution found. As exceeds 4% but UR exceeds target value.";
                                        break;
                                    }
                                    else
                                    {
                                        // Decrease rebar
                                        (double nextDia, double nextSpacing) = rebarHeirarchy.GetPreviousValue(dia, spacing);
                                        if (double.IsNaN(nextDia)) { optimisationStatus[rowNum] = "Error: No solution found. Reached the end of rebar configuration"; }
                                        diaRange.Value2 = nextDia;
                                        spacingRange.Value2 = nextSpacing;
                                        optimisationStatus[rowNum] = "Reinforcement decreased";
                                        rebarDecreased = true;
                                    }
                                }
                                else
                                {
                                    if (ur <= targetUr)
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        // Increase rebar
                                        (double nextDia, double nextSpacing) = rebarHeirarchy.GetNextValue(dia, spacing);
                                        if (double.IsNaN(nextDia)) { optimisationStatus[rowNum] = "Error: No solution found. Reached the end of rebar configuration."; }

                                        diaRange.Value2 = nextDia;
                                        spacingRange.Value2 = nextSpacing;
                                        optimisationStatus[rowNum] = "Reinforcement increased";
                                        if (rebarDecreased)
                                        {
                                            optimisationStatus[rowNum] = "Error: No solution found. As exceeds 4% but UR exceeds target value";
                                            break;
                                        }
                                    }
                                }
                                counter++;
                            }
                            if (counter >= rebarHeirarchy.table.GetLength(0)) { optimisationStatus[rowNum] = "Error: Maximum number of iterations met for rebar design"; }
                        }
                        catch (Exception ex)
                        {
                            optimisationStatus[rowNum] = "Error: Unexpected error encountered during rebar optimisation\n" + ex.Message;
                        }
                    }
                    #endregion
                }
                #endregion

                #region Write Status
                WriteToExcelRangeAsCol(checkSheetStatusRange, 0, 2, false, optimisationStatus);
                #endregion
            }
            catch (Exception ex) { throw new Exception("Error solving for rebar\n" + ex.Message); }
        }

        private List<TrackedRange> CreateDesignTrackRanges()
        {
            List<TrackedRange> trackedRanges = new List<TrackedRange>();

            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
            outputRange = outputRange.Worksheet.Cells[pierLabelRange.Row, outputRange.Column];

            Range rebarDiaRange = GetColRangeFromRanges(pierLabelRange, outputRange);
            TrackedRange trackedRange = new TrackedRange(rebarDiaRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);

            Range rebarSpacingRange = GetColRangeFromRanges(pierLabelRange, outputRange, 0, 1);
            trackedRange = new TrackedRange(rebarSpacingRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);

            return trackedRanges;
        }
        private void HighlightDesignChanges(List<TrackedRange> trackedRanges)
        {
            HighlightChangesForTrackedRanges(trackedRanges);

            #region Check for Errors
            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            Range designStatusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();
            designStatusRange = designStatusRange.Worksheet.Cells[pierLabelRange.Row, designStatusRange.Column + 2];
            designStatusRange = designStatusRange.Worksheet.Range[designStatusRange, designStatusRange.Offset[pierLabelRange.Rows.Count - 1]];

            HighlightErrors(designStatusRange, true);
            
            if (overwriteRebarCheck.Checked)
            {
                Range matchStatusRange = designStatusRange.Offset[0, -2];
                HighlightErrors(matchStatusRange, true);
            }

            if (unifyChangesCheck.Checked)
            {
                Range unifyStatusRange = designStatusRange.Offset[0, 2];
                HighlightErrors(unifyStatusRange, true);
            }
            #endregion
        }
        #endregion

        #region Unify Changes
        private void unifyChangesButt_Click(object sender, EventArgs e)
        {
            List<TrackedRange> trackedRanges = null;
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Stopwatch totalStopwatch = Stopwatch.StartNew();
                CreatePierLabelBackupSheet();
                CreateRebarTableBackupSheet();
                ReadStoreyTable();
                ReadRebarHeirarchy();
                ReadRebarTable();

                trackedRanges = CreateUnifyTrackRanges();
                UnifyChanges();

                Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                pierLabelRange.Worksheet.Activate();

                totalStopwatch.Stop();
                MessageBox.Show($"Total Execution Time: {totalStopwatch.ElapsedMilliseconds} ms", "Completed Unify Rebar");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                HighlightChangesForUnifyButton(trackedRanges);
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                rebarHeirarchy = null;
                rebarDic = null;
                storeyTracker = null;
            }
        }

        private List<TrackedRange> CreateUnifyTrackRanges()
        {
            List<TrackedRange> trackedRanges = new List<TrackedRange>();
            TrackedRange trackedRange = new TrackedRange(((RangeTextBox)TextBoxAttributeDic["rebarTable_WD"]).GetRangeFromFullAddress(), resetFontColourRebarTableCheck.Checked);
            trackedRanges.Add(trackedRange);

            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
            outputRange = outputRange.Worksheet.Cells[pierLabelRange.Row, outputRange.Column];

            Range rebarDiaRange = GetColRangeFromRanges(pierLabelRange, outputRange);
            Range rebarSpacingRange = GetColRangeFromRanges(pierLabelRange, outputRange, 0, 1);

            trackedRange = new TrackedRange(rebarDiaRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);
            trackedRange = new TrackedRange(rebarSpacingRange, resetFontColourCheckSheetCheck.Checked);
            trackedRanges.Add(trackedRange);

            return trackedRanges;
        }

        private void HighlightChangesForUnifyButton(List<TrackedRange> trackedRanges)
        {
            HighlightChangesForTrackedRanges(trackedRanges);
            Range statusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();
            Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            Range unifyStatusRange = GetColRangeFromRanges(pierLabelRange, statusRange, 0, 4);
            HighlightErrors(unifyStatusRange, true);
        }
        
        private void UnifyChanges()
        {
            try
            {
                #region Get Ranges
                Range pierLabelRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
                Range storeyRange = ((RangeTextBox)TextBoxAttributeDic["matchStoreyRange_WD"]).GetRangeFromFullAddress();
                storeyRange = storeyRange.Worksheet.Cells[pierLabelRange.Row, storeyRange.Column];
                Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
                outputRange = outputRange.Worksheet.Cells[pierLabelRange.Row, outputRange.Column];
                Range statusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();

                
                Range rebarDiaRange = GetColRangeFromRanges(pierLabelRange, outputRange);
                Range rebarSpacingRange = GetColRangeFromRanges(pierLabelRange, outputRange, 0, 1);

                double[] finalRebarDia = GetContentsAsDoubleArray(rebarDiaRange);
                double[] finalRebarSpacing = GetContentsAsDoubleArray(rebarSpacingRange);

                Range unifyStatusRange = GetColRangeFromRanges(pierLabelRange, statusRange, 0, 4);
                string[] unifyStatus = new string[unifyStatusRange.Rows.Count];
                #endregion

                #region Loop through data to update WallRebar
                for (int rowNum = 0; rowNum < pierLabelRange.Rows.Count; rowNum++)
                {
                    try
                    {
                        #region Get Excel Data
                        string pierLabel = pierLabelRange[rowNum + 1].Value2;
                        string etabsStoreyName = storeyRange.Offset[rowNum].Value2.ToString();
                        int targetStoreyNum = storeyTracker.GetStoreyNum(etabsStoreyName, "etabs");

                        Range finalRebarRange = outputRange.Offset[rowNum];
                        double finalDia = finalRebarRange.Offset[0, 0].Value2;
                        double finalSpacing = finalRebarRange.Offset[0, 1].Value2;
                        #endregion

                        #region Update Entry
                        if (!rebarDic.ContainsKey(pierLabel))
                        {
                            unifyStatus[rowNum] = $"Error: Pier label not found in rebar table. Skipped";
                            continue;
                        }
                        AssignedWallRebar assignedWallRebar = rebarDic[pierLabel];
                        WallRebarEntry wallRebarEntry = assignedWallRebar.GetEntry(targetStoreyNum);
                        wallRebarEntry.TryUpdateEntry(finalDia, finalSpacing, rebarHeirarchy);
                        unifyStatus[rowNum] = $"Ok";
                    }
                    catch (Exception ex) { unifyStatus[rowNum] = $"Error: {ex.Message}"; }
                    #endregion
                }
                #endregion

                #region Create and Print Final Rebar Table
                Range rebarTableRange = ((RangeTextBox)TextBoxAttributeDic["rebarTable_WD"]).GetRangeFromFullAddress();
                object[,] finalRebarTable = new object[rebarTableRange.Rows.Count, rebarTableRange.Columns.Count];
                int firstRowNum = rebarTableRange.Row; // rebar entry contains actual row number. Provide firstRowNum of rebar table so it can be offset to start at 0
                foreach (AssignedWallRebar assignedWallRebar in rebarDic.Values)
                {
                    assignedWallRebar.GetResults(ref finalRebarTable, firstRowNum);
                }
                
                WriteObjectToExcelRange(rebarTableRange, 0, 0, false, finalRebarTable);
                WriteToExcelRangeAsCol(unifyStatusRange, 0, 0, false, unifyStatus);
                #endregion

                #region Match Rebars Again
                ReadRebarTable();
                MatchRebars();
                #endregion
            }
            catch (Exception ex) { throw new Exception("Error unifying rebars\n" + ex.Message); }
        }
        #endregion

        #region Decomposer
        private void decomposeTable_Click(object sender, EventArgs e)
        {
            try
            {
                Range targetRange = GetFullTableRange();
                object[,] printObject = DecomposeTable(targetRange);

                (Worksheet outputSheet, Range outputRange) = ReplicateSheet(targetRange);
                ClearRangeForPrintingObject(outputRange, 0, 0, printObject);
                WriteObjectToExcelRange(outputRange, 0, 0, false, printObject);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private Range GetFullTableRange()
        {
            Range targetRange = ((RangeTextBox)TextBoxAttributeDic["decomposeRange_WD"]).GetRangeFromFullAddress();
            Range lastUsedCell = GetLastCell(targetRange.Worksheet, 1);

            Range startCell = targetRange.Worksheet.Cells[targetRange.Row, 1];
            Range endCell = targetRange.Worksheet.Cells[lastUsedCell.Row, targetRange.Column + targetRange.Columns.Count - 1];
            return targetRange.Worksheet.Range[startCell,endCell];
        }
        
        private (Worksheet outputSheet,Range outputRange) ReplicateSheet(Range targetRange)
        {
            Worksheet worksheet = targetRange.Worksheet;
            Worksheet outputSheet = CopyNewSheetAtBack(worksheet, worksheet.Name + "_decomposed", true);
            string targetRangeAddress = targetRange.Address;
            Range outputRange = outputSheet.Range[targetRangeAddress];
            return (outputSheet, outputRange);
        }
        
        private object[,] DecomposeTable(Range targetRange)
        {
            #region Loop through data
            object[,] dataContents = GetContentsAsObject2DArray(targetRange);
            List<object[]> finalOutput = new List<object[]>();
            int numCol = targetRange.Columns.Count;

            for (int rowNum = 0; rowNum < dataContents.GetLength(0); rowNum++)
            {
                if (dataContents[rowNum, 0] == null) { continue; }

                #region Decompose Pier Labels to parts
                string[] pierLabelsInMerged = SplitAndTrim(dataContents[rowNum, 0].ToString());
                #endregion
                #region Define start and end row number
                int startRowNum = rowNum;
                int endRowNum = rowNum;
                while ((endRowNum < dataContents.GetLength(0) - 1))
                {
                    if (dataContents[endRowNum + 1, 0] != null) { break; }
                    endRowNum++;
                }
                rowNum = endRowNum;
                #endregion

                #region Loop through to add to final arrage
                foreach (string pierLabel in pierLabelsInMerged)
                {
                    for (int localRowNum = startRowNum; localRowNum <= endRowNum; localRowNum++)
                    {
                        object[] rowData = new object[numCol];
                        rowData[0] = pierLabel;
                        for (int colNum = 1; colNum < numCol; colNum++)
                        {
                            rowData[colNum] = dataContents[localRowNum, colNum];
                        }
                        finalOutput.Add(rowData);
                    }
                }
                #endregion
            }
            #endregion

            #region Convert to object
            object[,] printObject = new object[finalOutput.Count, numCol];
            int printRowNum = 0;
            foreach (object[] row in finalOutput)
            {
                for (int colNum = 0; colNum < printObject.GetLength(1); colNum++)
                {
                    printObject[printRowNum, colNum] = row[colNum];
                }
                printRowNum++;
            }
            #endregion

            return printObject;
        }
        #endregion
    }

    #region Wall Rebar
    public class AssignedWallRebar
    {
        // Contains rebar information for a single pier label in the rebar table
        #region Init
        public string pierLabel;
        public Dictionary<int, WallRebarEntry> tableContents = new Dictionary<int, WallRebarEntry>(); // initial data in dictionary format, storey index points to data
        public List<string[]> tableContentsList = new List<string[]>(); // initial data in table format

        StoreyTracker storeyTracker;
        RebarHeirarchy rebarHeirarchy;
        public AssignedWallRebar(string name, StoreyTracker storeyTracker)
        {
            // Used for matching rebars only
            pierLabel = name;
            this.storeyTracker = storeyTracker;
        }
        public AssignedWallRebar(string name, StoreyTracker storeyTracker, RebarHeirarchy rebarHeirarchy)
        {
            // Used for designing rebar, rebar heirarchy is required
            pierLabel = name;
            this.storeyTracker = storeyTracker;
            this.rebarHeirarchy = rebarHeirarchy;
        }

        public void AddRow(Range row)
        {
            WallRebarEntry wallRebarEntry = new WallRebarEntry(row, storeyTracker, rebarHeirarchy);
            tableContents.Add(wallRebarEntry.StartStoreyNum, wallRebarEntry);
        }
        #endregion

        #region Modify Data
        public void TryUpdateEntry(int storeyNum, double newDia, double newSpacing)
        {
            WallRebarEntry wallRebarEntry = GetEntry(storeyNum);
            wallRebarEntry.TryUpdateEntry(newDia, newSpacing, rebarHeirarchy);
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
                WallRebarEntry wallRebarEntry = tableContents[startStoreyNumSorted[i]];
                endStoreyNumSorted[i] = wallRebarEntry.EndStoreyNum;
            }
        }
        public object[] GetStoreyData(int targetStoreyNum)
        { 
            WallRebarEntry wallRebarEntry = GetEntry(targetStoreyNum);
            return wallRebarEntry.originalRowContents;
        }

        public WallRebarEntry GetEntry(int targetStoreyNum)
        {
            for (int i = 0; i < startStoreyNumSorted.Length; i++)
            {
                if (targetStoreyNum >= startStoreyNumSorted[i] && targetStoreyNum <= endStoreyNumSorted[i])
                {
                    WallRebarEntry wallRebarEntry = tableContents[startStoreyNumSorted[i]];
                    return wallRebarEntry;
                }
            }

            throw new Exception("Unable to find target storey");
        }
        #endregion

        #region Fill Object for print
        public void GetResults(ref object[,] results, int firstRowNum)
        {
            foreach (WallRebarEntry wallRebarEntry in tableContents.Values)
            {
                wallRebarEntry.GetResults(ref results, firstRowNum);
            }
        }
        #endregion

    }

    public class WallRebarEntry
    {
        // Contains rebar information for a single storey for a pier label in the rebar table
        #region Init
        public object[] originalRowContents;
        public int excelRowNumber;
        bool modified = false;

        StoreyTracker storeyTracker;
        
        double initialDia;
        double initialSpacing;

        public WallRebarEntry(Range row, StoreyTracker storeyTracker, RebarHeirarchy rebarHeirarchy)
        {
            this.storeyTracker = storeyTracker;

            originalRowContents = GetContentsAsObject1DArray(row); // Consider making this an object array instead
            excelRowNumber = row.Row;

            #region Check entries are double
            for (int i = 3; i < originalRowContents.Length; i++)
            {
                if (!(originalRowContents[i] is double))
                {
                    Range errorCell = row.Cells[i + 1];
                    string address = errorCell.Address;
                    throw new ArgumentException($"Cell {errorCell.Worksheet.Name}!{errorCell.Address[false, false]} does not contain a number");
                }
            }

            initialDia = (double)originalRowContents[3];
            initialSpacing = (double)originalRowContents[4];
            #endregion


            // Skip adding currentHeirarchyRowNum if we are only matching (no rebarHeirarchy defined)
            if (rebarHeirarchy != null) { currentHeirarchyRowNum = rebarHeirarchy.GetPosition(initialDia, initialSpacing); }
            
        }
        #endregion

        #region Update Value
        public object[] finalRowContents = null;
        int currentHeirarchyRowNum;
        public void TryUpdateEntry(double newDia, double newSpacing, RebarHeirarchy rebarHeirarchy)
        {
            int newHeirarchyRowNum = rebarHeirarchy.GetPosition(newDia, newSpacing);

            if (newHeirarchyRowNum > currentHeirarchyRowNum)
            {
                if (!modified)
                {
                    finalRowContents = new object[originalRowContents.Length];
                    originalRowContents.CopyTo(finalRowContents, 0);
                    modified = true;
                }

                currentHeirarchyRowNum = newHeirarchyRowNum;
                finalRowContents[3] = newDia.ToString();
                finalRowContents[4] = newSpacing.ToString();
            }
        }
        #endregion

        #region Get Sets
        public int StartStoreyNum
        {
            get
            {
                string startStoreyName = originalRowContents[1].ToString();
                return storeyTracker.GetStoreyNum(startStoreyName, "design");
            }
        }
        public int EndStoreyNum
        {
            get
            {
                string endStoreyName = originalRowContents[2].ToString();
                return storeyTracker.GetStoreyNum(endStoreyName, "design");
            }
        }
        #endregion

        #region MyRegion
        internal void GetResults(ref object[,] results, int firstRowNum)
        {
            // Get Row Number
            int relativeRowNum = excelRowNumber - firstRowNum;
            if (relativeRowNum < 0) { throw new ArgumentException($"Error filling table for excel row number {excelRowNumber}, relative row number < 0"); }
            // Update entry
            if (!modified)
            {
                for (int colNum  = 0; colNum < originalRowContents.Length; colNum++)
                {
                    results[relativeRowNum,colNum] = originalRowContents[colNum];
                }
            }
            else
            {
                for (int colNum = 0; colNum < originalRowContents.Length; colNum++)
                {
                    results[relativeRowNum, colNum] = finalRowContents[colNum];
                }
            }
        }
        #endregion
    }
    #endregion

    #region Reference Tables
    public class StoreyTracker
    {
        Dictionary<string, int> etabsStoreyDicNameToNum = new Dictionary<string, int>();
        Dictionary<string, int> designStoreyDicNameToNum = new Dictionary<string, int>();
        Dictionary<int, string> etabsStoreyDicNumToName = new Dictionary<int, string>();
        Dictionary<int, string> designStoreyDicNumToName = new Dictionary<int, string>();
        public StoreyTracker(Range storeyTable)
        {
            GetStoreyTable(storeyTable);
        }
        private void GetStoreyTable(Range storeyTable)
        {
            foreach (Range range in storeyTable.Rows)
            {
                int storeyNumber = Int32.Parse(range.Cells[1].Value2.ToString());
                string etabsStoreyName = range.Cells[2].Value2.ToString();
                string designStoreyName = range.Cells[3].Value2.ToString();

                etabsStoreyDicNameToNum.Add(etabsStoreyName, storeyNumber);
                designStoreyDicNameToNum.Add(designStoreyName, storeyNumber);
                etabsStoreyDicNumToName.Add(storeyNumber, etabsStoreyName);
                designStoreyDicNumToName.Add(storeyNumber, designStoreyName);
            }
        }

        public int GetStoreyNum(string name, string type)
        {
            if (type == "etabs")
            {
                if (!etabsStoreyDicNameToNum.ContainsKey(name)) { throw new Exception($"ETABS Storey Name \"{name}\" not found in reference table."); }
                return etabsStoreyDicNameToNum[name];
            }
            else if (type == "design")
            {
                if (!designStoreyDicNameToNum.ContainsKey(name)) { throw new Exception($"Design Storey Name \"{name}\" not found in reference table."); }
                return designStoreyDicNameToNum[name];
            }
            else
            {
                throw new Exception($"Invalid type {type} for GetStoreyNum");
            }
        }

        public string GetStoreyName(int num, string type)
        {
            if (type == "etabs")
            {
                if (!etabsStoreyDicNumToName.ContainsKey(num)) { throw new Exception("ETABS Storey Number not found in reference table."); }
                return etabsStoreyDicNumToName[num];
            }
            else if (type == "design")
            {
                if (!designStoreyDicNumToName.ContainsKey(num)) { throw new Exception("Design Storey Number not found in reference table."); }
                return designStoreyDicNumToName[num];
            }
            else
            {
                throw new Exception($"Invalid type {type} for GetStoreyNum");
            }
        }
    }

    public class RebarHeirarchy
    {
        #region Init
        public object[,] table;
        Dictionary<string, int> rebarToPosition = new Dictionary<string, int>();
        public RebarHeirarchy(Range rebarHeirarchyRange)
        {
            CreateObjectTable(rebarHeirarchyRange);
            CreateHashTable();
        }
        private void CreateObjectTable(Range rebarHeirarchyRange)
        {
            table = new object[rebarHeirarchyRange.Rows.Count, rebarHeirarchyRange.Columns.Count];

            for (int rowNum = 0; rowNum < rebarHeirarchyRange.Rows.Count; rowNum++)
            {
                for (int colNum = 0; colNum < rebarHeirarchyRange.Columns.Count; colNum++)
                {
                    Range cell = rebarHeirarchyRange.Cells[rowNum + 1, colNum + 1];
                    bool canParse = double.TryParse(cell.Value2.ToString(), out double cellValue);
                    if (!canParse) { throw new ArgumentException($"Unable to convert value \"{cell.Value2.ToString()}\" at {cell.Worksheet.Name}!{cell.Address[false, false]} for Rebar Heirarchy"); }
                    table[rowNum, colNum] = double.Parse(cell.Value2.ToString());
                }
            }
        }
        private void CreateHashTable()
        {
            for (int rowNum = 0; rowNum < table.GetLength(0); rowNum++)
            {
                string key = ConvertRebarSpacingToKey((double)table[rowNum, 0], (double)table[rowNum, 1]);
                rebarToPosition.Add(key, rowNum);
            }
        }
        #endregion

        #region Get Values
        public (double nextDia, double nextSpacing) GetNextValue(double dia, double spacing)
        {
            int currentRowNum = GetPosition(dia, spacing);
            if (currentRowNum >= table.GetLength(0)) { return (double.NaN, double.NaN); }
            double nextDia = (double)table[currentRowNum + 1, 0];
            double nextSpacing = (double)table[currentRowNum + 1, 1];
            return (nextDia, nextSpacing);
        }

        public (double nextDia, double nextSpacing) GetPreviousValue(double dia, double spacing)
        {
            int currentRowNum = GetPosition(dia, spacing);
            if (currentRowNum == 0) { return (double.NaN, double.NaN); }
            double nextDia = (double)table[currentRowNum - 1, 0];
            double nextSpacing = (double)table[currentRowNum - 1, 1];
            return (nextDia, nextSpacing);
        }
        #endregion

        #region Key
        private string ConvertRebarSpacingToKey(double dia, double spacing)
        {
            string diaString = dia.ToString("#.####");
            string spacingString = spacing.ToString("#.####");
            string key = diaString + "-" + spacingString;
            return key;
        }
        #endregion

        #region Get Position
        public int GetPosition(double dia, double spacing)
        {
            string currentKey = ConvertRebarSpacingToKey(dia, spacing);
            return rebarToPosition[currentKey];
        }
        #endregion
    }
    #endregion

    #region Highlight Changes
    public class TrackedRange
    {
        object[,] originalContents;
        Range range;
        public TrackedRange(Range range, bool resetRange = true)
        {
            this.range = range;
            originalContents = GetContentsAsObject2DArray(range);
            if (resetRange) { ResetRange(); }
        }
        public void ResetRange()
        {
            range.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
        }
        public void HighlightChanges()
        {
            for (int rowNum = 0; rowNum < originalContents.GetLength(0); rowNum++)
            {
                for (int colNum = 0; colNum < originalContents.GetLength(1); colNum++)
                {
                    object originalValue = originalContents[rowNum, colNum];

                    Range cell = range.Cells[rowNum + 1, colNum + 1];
                    object currentValue = cell.Value2;

                    bool format = false;
                    if (originalValue is string originalString && currentValue is string currentString)
                    {
                        if (originalString != currentString) { format = true; }
                    }
                    else if (originalValue is double originalDouble && currentValue is double currentDouble)
                    {
                        if (!originalDouble.Equals(currentDouble)) { format = true; }
                    }
                    // Handle other types as needed
                    else if (!Equals(originalValue, currentValue)) { format = true; }

                    if (format) { cell.Font.Color = Color.Red; }
                }
            }
        }
    }
    #endregion
}
