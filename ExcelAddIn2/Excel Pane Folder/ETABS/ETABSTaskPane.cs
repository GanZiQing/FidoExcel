using ETABSv1;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using MigraDoc.DocumentObjectModel;
using PdfSharp.Snippets.Font;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;
using static ExcelAddIn2.EtabsFunctions;

namespace ExcelAddIn2
{
    public partial class ETABSTaskPane : UserControl
    {
        #region Init
        public ETABSTaskPane()
        {
            InitializeComponent();
            CreateAttributes();
            AddHeaders();
            AddToolTips();
        }

        Dictionary<string, object> attributeDic = new Dictionary<string, object>();
        private void CreateAttributes()
        {
            AttributeTextBox thisTBAtt = new RangeTextBox("storyRange_AWL", dispStoryRange, setStoryRange);
            attributeDic.Add(thisTBAtt.attName, thisTBAtt);

            thisTBAtt = new RangeTextBox("jointDataRange_AWL", dispJointDataRange, setJointDataRange);
            attributeDic.Add(thisTBAtt.attName, thisTBAtt);

            CustomAttribute thisAtt = new ComboBoxAttribute("storySortOrder_AWL", dispStorySortOrder, "Top to Bottom");
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new ComboBoxAttribute("jointSortOrder_AWL", dispJointSortOrder, "Z, X, Y");
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new ComboBoxAttribute("windLoadDir_AWL", dispWindLoadDir, "X");
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new CheckBoxAttribute("replaceLoad_AWL", replaceLoadCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new CheckBoxAttribute("refreshView_AWL", refreshViewCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);

            CreateAttributesForBaseShear();
            CreateAttributesForUtilities();
        }

        private void AddHeaders()
        {
            List<string> headers;
            #region Get ETABS 
            headers = new List<string>
            {
                "Story Name",
                "Story Elevation [m]",
                "Effective Height [m]"
            };
            AddHeaderMenuToButton(getStoryData, headers);

            headers = new List<string>
            {
                "Joint UN",
                "X",
                "Y",
                "Z"
            };
            AddHeaderMenuToButton(getJointCoordinates, headers);
            #endregion

            #region Calculate WL
            headers = new List<string>
            {
                "Story Name",
                "Story Elevation [m]",
                "Effective Height [m]",
                "Start WL [kN/m]",
                "End WL [kN/m]",
                "Min X [m]",
                "Max X [m]",
                "Min Y [m]",
                "Max Y [m]"
            };
            AddHeaderMenuToButton(setStoryRange, headers);

            #endregion

            #region Assign WL
            headers = new List<string>
            {
                "Joint UN",
                "X [m]",
                "Y [m]",
                "Z [m]",
                "Start Coord [m]",
                "End Coord [m]",
                "Eff Width [m]",
                "Start WL [kN/m]",
                "End WL [kN/m]",
                "Total WL [kN]",
                "Direction",
                "Load Pattern Name",
                "Status"
            };
            AddHeaderMenuToButton(setJointDataRange, headers);
            //AddHeaderMenuToButton(calAWL, headers);
            //AddHeaderMenuToButton(assignWL, headers);
            #endregion

            #region Error Joints
            headers = new List<string>
            {
                "Joint Label",
                "Joint UN",
                "X [m]",
                "Y [m]",
                "Z [m]",
                "Storey Name",
                "Unstable DOF",
                "Status"
            };
            AddHeaderMenuToButton(groupAndImportLog, headers);
            #endregion
        }

        private void AddToolTips()
        {
            #region Get ETABS
            toolTip1.SetToolTip(getStoryData,
                "Gets selected joint info for attached instance of ETABS\n" +
                "  Story Name\n" +
                "  Story Elevation\n" +
                "  Effective Height (calculated assuming 1st sty elevation = 0)\n"
                );

            toolTip1.SetToolTip(getJointCoordinates,
                "Gets selected joint info for attached instance of ETABS, values rounded to nearest 4dp\n" +
                "  Joint UN\n" +
                "  X\n" +
                "  Y\n" +
                "  Z\n"
                );

            toolTip1.SetToolTip(getLoadPatterns,
                "Gets all load patterns currently defined in attached instance of ETABS"
                );
            #endregion

            #region Calculate WL
            toolTip1.SetToolTip(setStoryRange,
                "Takes input in the following order:\n" +
                "  Story Name\n" +
                "  Story Elevation [m]\n" +
                "  Effective Height [m]\n" +
                "  Minimum WL Value [kN/m]\n" +
                "  Maximum WL Value [kN/m]\n" +
                "  Min X [m]\n" +
                "  Max X [m]\n" +
                "  Min Y [m]\n" +
                "  Max Y [m]"
                );

            toolTip1.SetToolTip(calAWL,
                "Calculates wind load based on:\n" +
                "  Data in Story Range\n" +
                "  Currently selected joints in attached instance of ETABS"
                );
            #endregion

            #region Assign WL
            toolTip1.SetToolTip(setJointDataRange,
                "Takes input in the following order:\n" +
                "  Joint UN\n" +
                "  X [m]\n" +
                "  Y [m]\n" +
                "  Z [m]\n" +
                "  Start Coord [m]\n" +
                "  End Coord [m]\n" +
                "  Eff Width [m]\n" +
                "  Start WL [kN/m]\n" +
                "  End WL [kN/m]\n" +
                "  Total WL [kN]\n" +
                "  Direction\n" +
                "  Load Pattern Name\n" +
                "  Status\n" +
                "Only Joint UN, Total WL, Direction, Load Pattern is read."
                );

            toolTip1.SetToolTip(assignWL,
                "Assigns wind load based on the data provided in Joint Data Range"
                );

            toolTip1.SetToolTip(replaceLoadCheck,
                "If checked, this will replace the entire load pattern (including other directions) with current loading.\n" +
                "If unchecked, current loading will be added to existing loading\n" +
                "Does not affect other load patterns"
                );
            #endregion

        }
        #endregion

        #region AWL
        #region Get ETABS Data
        private void getStoryData_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out ETABSv1.cOAPI etabsObject, out ETABSv1.cSapModel sapModel, true);

                #region Get Data Using ETABS API and Redorder
                int ret = 0;
                double BaseElevation = 0;
                int NumberStories = 0;
                string[] storyNames = new string[0];
                double[] storyElevations = new double[0];
                double[] storyHeights = new double[0];
                bool[] isMasterStory = new bool[0];
                string[] similarToStory = new string[0];
                bool[] spliceAbove = new bool[0];
                double[] spliceHeight = new double[0];
                int[] color = new int[0];

                ret = sapModel.Story.GetStories_2(ref BaseElevation, ref NumberStories, ref storyNames, ref storyElevations, ref storyHeights, ref isMasterStory, ref similarToStory, ref spliceAbove, ref spliceHeight, ref color);
                if (ret != 0) { throw new Exception("Unable to get story info"); }
                #endregion

                #region Calculate Effective Height
                double[] effHeight = new double[storyNames.Length];
                // Calculate for first value
                if (storyElevations[0] <= 0) { effHeight[0] = 0; }
                else
                {
                    int i = 0;
                    effHeight[i] = storyHeights[i] / 2 + storyHeights[i + 1] / 2;
                }

                // Calculate for mid
                for (int i = 1; i < storyElevations.Length - 1; i++)
                {
                    if (storyElevations[i] <= 0) { effHeight[i] = 0; continue; }
                    effHeight[i] = storyHeights[i] / 2 + storyHeights[i + 1] / 2;
                }

                // Calculate for last value
                effHeight[storyNames.Length - 1] = storyHeights[storyNames.Length - 1] / 2;
                #endregion

                #region Sort Data and Add Base Story
                string sortType = ((ComboBoxAttribute)attributeDic["storySortOrder_AWL"]).attValue;

                string[] storyNamesPrint = new string[storyNames.Length + 1];
                double[] storyElevationsPrint = new double[storyNames.Length + 1];
                double[] effHeightPrint = new double[storyNames.Length + 1];

                Array.Copy(storyNames, 0, storyNamesPrint, 1, storyNames.Length);
                Array.Copy(storyElevations, 0, storyElevationsPrint, 1, storyNames.Length);
                Array.Copy(effHeight, 0, effHeightPrint, 1, storyNames.Length);
                storyNamesPrint[0] = "Base";
                storyElevationsPrint[0] = BaseElevation;
                effHeightPrint[0] = 0;

                if (sortType == "Top to Bottom")
                {
                    Array.Reverse(storyNamesPrint);
                    Array.Reverse(storyElevationsPrint);
                    Array.Reverse(effHeightPrint);
                }
                else if (sortType == "Bottom to Top") { }
                else { throw new NotImplementedException($"Sort type \"{sortType}\" not implemented"); }
                #endregion

                #region Write to Excel
                WriteToExcelRangeAsCol(null, 0, 0, true, storyNamesPrint, storyElevationsPrint, effHeightPrint);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getJointCoordinates_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out ETABSv1.cOAPI etabsObject, out ETABSv1.cSapModel sapModel, true);

                string sortType = ((ComboBoxAttribute)attributeDic["jointSortOrder_AWL"]).attValue;
                (string[] selectedJoints, double[] Xs, double[] Ys, double[] Zs) = GetSortedJoints(sapModel, sortType);

                WriteToExcelRangeAsCol(null, 0, 0, true, selectedJoints, Xs, Ys, Zs);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getLoadPatterns_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out ETABSv1.cOAPI etabsObject, out ETABSv1.cSapModel sapModel, true);
                int NumberNames = 0;
                string[] MyName = new string[0];
                int ret = sapModel.LoadPatterns.GetNameList(ref NumberNames, ref MyName);

                WriteToExcelRangeAsCol(null, 0, 0, true, MyName);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region Asymmetrical Wind Load

        ETABSv1.cOAPI etabsObject;
        ETABSv1.cSapModel sapModel;
        private void calAWL_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out etabsObject, out sapModel, true);

                StoryTable storyTable = ReadstoryTable();

                CalculateAWL(storyTable);

                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            finally
            {
                etabsObject = null;
                sapModel = null;
            }
        }

        private StoryTable ReadstoryTable()
        {
            Range storyRange = ((RangeTextBox)attributeDic["storyRange_AWL"]).GetRangeFromFullAddress();
            StoryTable storyTable = new StoryTable(storyRange);
            return storyTable;
        }

        private void CalculateAWL(StoryTable storyTable)
        {
            #region Get sorted joints
            string wlDir = ((ComboBoxAttribute)attributeDic["windLoadDir_AWL"]).attValue;
            string[] selectedJoints;
            double[] Xs;
            double[] Ys;
            double[] Zs;
            if (wlDir == "X") { (selectedJoints, Xs, Ys, Zs) = GetSortedJoints(sapModel, "Z, Y, X"); }
            else if (wlDir == "Y") { (selectedJoints, Xs, Ys, Zs) = GetSortedJoints(sapModel, "Z, X, Y"); }
            else { throw new Exception($"Invalide wind load direction \"{wlDir}\""); }
            #endregion

            #region Group sorted joints into Each Floor
            Dictionary<string, List<int>> elevationToJointIndex = new Dictionary<string, List<int>>();
            for (int i = 0; i < selectedJoints.Length; i++)
            {
                string elevationString = Zs[i].ToString("#.####");
                if (elevationString == "") { elevationString = "0"; }
                if (!elevationToJointIndex.ContainsKey(elevationString)) { elevationToJointIndex.Add(elevationString, new List<int>()); }
                List<int> elevationIndexList = elevationToJointIndex[elevationString];
                elevationIndexList.Add(i);
            }
            #endregion

            #region Calculate for each Floor
            string[] status = new string[selectedJoints.Length];
            double[] startCoord = new double[selectedJoints.Length];
            double[] endCoord = new double[selectedJoints.Length];
            double[] effWidth = new double[selectedJoints.Length];
            double[] startWL = new double[selectedJoints.Length];
            double[] endWL = new double[selectedJoints.Length];
            double[] windLoad = new double[selectedJoints.Length];
            string direction = ((ComboBoxAttribute)attributeDic["windLoadDir_AWL"]).attValue;
            foreach (KeyValuePair<string, List<int>> entry in elevationToJointIndex)
            {
                CalculateAWLForOneStory(storyTable, direction, entry.Key, entry.Value,
                    selectedJoints, Xs, Ys, Zs,
                    ref status, ref startCoord, ref endCoord, ref effWidth, ref startWL, ref endWL, ref windLoad);
            }
            string[] directionArray = new string[selectedJoints.Length];
            for (int i = 0; i < selectedJoints.Length; i++) { directionArray[i] = direction; }
            #endregion

            #region Check Overlap
            // Check Overlapping joint
            double[] refCoord;
            if (direction == "X") { refCoord = Ys; }
            else { refCoord = Xs; }
            for (int i = 0; i < refCoord.Length - 1; i++)
            {
                if (refCoord[i] == refCoord[i + 1])
                {
                    if (status[i + 1] != null) { status[i] += " "; }
                    status[i + 1] += "Warning, overlapping joints";
                    windLoad[i + 1] = 0;
                }
            }
            #endregion

            #region Write to Excel
            WriteToExcelRangeAsCol(null, 0, 0, true, selectedJoints, Xs, Ys, Zs, startCoord, endCoord, effWidth, startWL, endWL, windLoad, directionArray);
            WriteToExcelRangeAsCol(null, 0, 12, false, status);
            #endregion

            sapModel.View.RefreshView();
        }

        private void CalculateAWLForOneStory(StoryTable storyTable, string direction, string elevationString, List<int> jointIndexes,
            string[] selectedJoints, double[] xs, double[] ys, double[] zs,
            ref string[] status, ref double[] globalStartCoord, ref double[] globalEndCoord, ref double[] globalEffWidth, ref double[] globalStartWL, ref double[] globalEndWL, ref double[] globalWindLoad)
        {
            #region Checks
            double elevation = double.Parse(elevationString);
            double effHeight = storyTable.EffHeight(elevation);
            if (effHeight == 0) // No WL to calculate
            {
                foreach (int index in jointIndexes)
                {
                    status[index] = "Effective Height is 0";
                }
                return;
            }
            #endregion

            #region Define reference values
            double minWL = storyTable.MinWL(elevation);
            double maxWL = storyTable.MaxWL(elevation);
            double minCoord;
            double maxCoord;

            if (direction == "X")
            {
                minCoord = storyTable.MinY(elevation);
                maxCoord = storyTable.MaxY(elevation);
            }
            else if (direction == "Y")
            {
                minCoord = storyTable.MinX(elevation);
                maxCoord = storyTable.MaxX(elevation);
            }
            else { throw new Exception($"Direction {direction} is invalid."); }
            #endregion

            #region Create local Array
            List<double> validCoordList = new List<double>();
            List<int> validIndexsList = new List<int>();
            foreach (int index in jointIndexes)
            {
                if (direction == "Y")
                {
                    if (xs[index] < minCoord) { status[index] = "No WL, position is smaller than min X value"; continue; }
                    if (xs[index] > maxCoord) { status[index] = "No WL, position is greater than max X value"; continue; }
                    validCoordList.Add(xs[index]);
                }
                else if (direction == "X")
                {
                    if (ys[index] < minCoord) { status[index] = "No WL, position is smaller than min X value"; continue; }
                    if (ys[index] > maxCoord) { status[index] = "No WL, position is greater than max X value"; continue; }
                    validCoordList.Add(ys[index]);
                }

                validIndexsList.Add(index);
            }
            double[] validCoords = validCoordList.ToArray();
            int[] validIndexes = validIndexsList.ToArray();
            #endregion

            #region Gatekeep if only 1 joint provided
            if (validCoords.Length == 0) { throw new Exception($"Error: No valid joints"); }
            else if (validCoords.Length == 1)
            {
                int globalIndex = validIndexes[0];
                globalStartCoord[globalIndex] = minCoord;
                globalEndCoord[globalIndex] = maxCoord;
                globalEffWidth[globalIndex] = maxCoord - minCoord;
                globalStartWL[globalIndex] = minWL;
                globalEndWL[globalIndex] = maxWL;
                globalWindLoad[globalIndex] = Math.Round(((minWL + maxWL) / 2) * globalEffWidth[globalIndex], 2);
                return;
            }

            #endregion

            #region Calculate Coordinates
            double[] localStartCoords = new double[validCoords.Length];
            double[] localEndCoords = new double[validCoords.Length];

            // Deal with first entry
            localStartCoords[0] = minCoord;
            localEndCoords[0] = (validCoords[1] + validCoords[0]) / 2;

            // Deal with typical entry
            for (int i = 1; i < validCoords.Length - 1; i++)
            {
                localStartCoords[i] = localEndCoords[i - 1];
                localEndCoords[i] = (validCoords[i] + validCoords[i + 1]) / 2;
            }

            // Deal with final entry
            localStartCoords[validCoords.Length - 1] = localEndCoords[validCoords.Length - 2];
            localEndCoords[validCoords.Length - 1] = maxCoord;
            #endregion

            #region Calculate WL
            double[] localEffWidth = new double[validCoords.Length];
            double[] localStartWL = new double[validCoords.Length];
            double[] localEndWL = new double[validCoords.Length];
            double[] localWL = new double[validCoords.Length];
            Func<double, double> windLoadEquation;
            if (direction == "X")
            {
                windLoadEquation = storyTable.WindLoadInX(elevation);
            }
            else if (direction == "Y")
            {
                windLoadEquation = storyTable.WindLoadInY(elevation);
            }
            else { throw new Exception($"Direction {direction} is invalid."); }

            for (int i = 0; i < validCoords.Length; i++)
            {
                localEffWidth[i] = localEndCoords[i] - localStartCoords[i];
                localStartWL[i] = windLoadEquation(localStartCoords[i]);
                localEndWL[i] = windLoadEquation(localEndCoords[i]);
                double avgWL = (localStartWL[i] + localEndWL[i]) / 2;
                localWL[i] = localEffWidth[i] * avgWL;
            }
            #endregion

            #region Map to global arrays
            for (int i = 0; i < validCoords.Length; i++)
            {
                int globalIndex = validIndexes[i];
                globalStartCoord[globalIndex] = localStartCoords[i];
                globalEndCoord[globalIndex] = localEndCoords[i];
                globalEffWidth[globalIndex] = localEffWidth[i];
                globalStartWL[globalIndex] = Math.Round(localStartWL[i], 2);
                globalEndWL[globalIndex] = Math.Round(localEndWL[i], 2);
                globalWindLoad[globalIndex] = Math.Round(localWL[i], 2);
            }
            #endregion
        }
        #endregion

        private void assignWL_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Info
                Range sourceRange = ((RangeTextBox)attributeDic["jointDataRange_AWL"]).GetRangeFromFullAddress();
                CheckRangeSize(sourceRange, 0, 13, "Joint Data Range", true);

                string[] UN = GetContentsAsStringArray(sourceRange.Columns[1], false);
                double[] WL = GetContentsAsDoubleArray(sourceRange.Columns[10]);
                string[] direction = GetContentsAsStringArray(sourceRange.Columns[11], false);
                string[] loadPatternName = GetContentsAsStringArray(sourceRange.Columns[12], false);
                string[] status = new string[UN.Length];
                #endregion

                #region Assign in ETABS
                InitializeETABS(out ETABSv1.cOAPI etabsObject, out ETABSv1.cSapModel sapModel, true);
                for (int i = 0; i < UN.Length; i++)
                {
                    double[] forces = new double[6];
                    // Check Direction
                    double wlValue = 0;

                    if (direction[i] == "X") { forces[0] = WL[i]; wlValue = WL[i]; }
                    else if (direction[i] == "Y") { forces[1] = WL[i]; wlValue = WL[i]; }
                    else { throw new Exception($"Direction {direction[i]} for UN {UN[i]}is invalid."); }

                    if (wlValue == 0)
                    {
                        if (replaceLoadCheck.Checked)
                        {
                            int ret = sapModel.PointObj.DeleteLoadForce(UN[i], loadPatternName[i]);
                            if (ret != 0) { status[i] = $"Error: Unable to delete joint forces"; }
                            else { status[i] = $"Joint forces deleted for {loadPatternName[i]}"; }
                        }
                        else { } // skip if load is 0
                    }
                    else
                    {
                        int ret = sapModel.PointObj.SetLoadForce(UN[i], loadPatternName[i], ref forces, replaceLoadCheck.Checked);
                        if (ret != 0) { status[i] = $"Error: Unable to set forces for joint"; }
                    }
                }
                #endregion

                WriteToExcelRangeAsCol(sourceRange, 0, 12, false, status);

                if (refreshViewCheck.Checked) { sapModel.View.RefreshView(); }
                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region Smart Replicate
        private void replicateByDispBut_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Info
                Range activeRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                CheckRangeSize(activeRange, 0, 3, "Offset Values");

                double[,] offsetValues = GetContentsAsDouble2DArray(activeRange);
                #endregion

                #region Calculate spacing

                //double[,] offsetValues;

                //if (absoluteDispCheck.Checked)
                //{
                //    offsetValues = excelValues;
                //}
                //else
                //{
                //    offsetValues = new double[excelValues.GetLength(0), excelValues.GetLength(1)];
                //    for (int i = 0; i < excelValues.GetLength(0); i++)
                //    {
                //        if (i == 0)
                //        {
                //            offsetValues[i, 0] = excelValues[i, 0];
                //            offsetValues[i, 1] = excelValues[i, 1];
                //            offsetValues[i, 2] = excelValues[i, 2];
                //            continue;
                //        }

                //        offsetValues[i, 0] = offsetValues[i - 1, 0] + excelValues[i, 0];
                //        offsetValues[i, 1] = offsetValues[i - 1, 1] + excelValues[i, 1];
                //        offsetValues[i, 2] = offsetValues[i - 1, 2] + excelValues[i, 2];
                //    }
                //}
                #endregion

                #region Init ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion

                #region Replicate Elements
                ReplicateSelectedElement(etabsObject, sapModel, offsetValues);
                #endregion

                sapModel.View.RefreshView();
                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void replicateBySpacingDisp_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Info
                Range activeRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                CheckRangeSize(activeRange, 0, 3, "Offset Values");

                double[,] excelValues = GetContentsAsDouble2DArray(activeRange);
                #endregion

                #region Calculate offset from spacing
                double[,] offsetValues;
                offsetValues = new double[excelValues.GetLength(0), excelValues.GetLength(1)];
                for (int i = 0; i < excelValues.GetLength(0); i++)
                {
                    if (i == 0)
                    {
                        offsetValues[i, 0] = excelValues[i, 0];
                        offsetValues[i, 1] = excelValues[i, 1];
                        offsetValues[i, 2] = excelValues[i, 2];
                        continue;
                    }

                    offsetValues[i, 0] = offsetValues[i - 1, 0] + excelValues[i, 0];
                    offsetValues[i, 1] = offsetValues[i - 1, 1] + excelValues[i, 1];
                    offsetValues[i, 2] = offsetValues[i - 1, 2] + excelValues[i, 2];
                }
                #endregion

                #region Init ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion

                #region Replicate Elements
                ReplicateSelectedElement(etabsObject, sapModel, offsetValues);
                #endregion

                sapModel.View.RefreshView();
                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region ETABS Error Analysis
        private string getAnalysisFile(cSapModel sapModel, string extension)
        {
            string cleanExtension = extension.StartsWith(".") ? extension : "." + extension;
            string etabsFilePath = sapModel.GetModelFilename();
            string returnFileName = Path.ChangeExtension(etabsFilePath, cleanExtension);
            return returnFileName;
        }
        private void groupAndImportLog_Click(object sender, EventArgs e)
        {
            // Copy from old code, to refractor later
            try
            {
                #region Init ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                // Get Storey Details
                (_, Dictionary<double, string> elevationToStoreyMap) = GetEtabsStoreys(sapModel);
                #endregion

                #region Create Lists
                string grpNm = ".E.Error Joints";
                Dictionary<string, EtabsJoint> unstableJoints = new Dictionary<string, EtabsJoint>();
                #endregion

                #region Analyse Files
                string logFilePath = getAnalysisFile(sapModel, ".LOG");
                int numLines = File.ReadLines(logFilePath).Count();


                using (StreamReader sr = new StreamReader(logFilePath))
                {
                    for (int i = 0; i < numLines; i++)
                    {
                        string line = sr.ReadLine();
                        if (line.Length > 6)
                        {
                            if (line.Substring(1, 5) == "Joint")
                            {
                                string[] textRowSplitted = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                                // Create or get joint object
                                EtabsJoint joint;
                                if (!unstableJoints.ContainsKey(textRowSplitted[1]))
                                {
                                    string uniqueName = textRowSplitted[1];
                                    joint = new EtabsJoint
                                    (
                                        uniqueName,
                                        "",
                                        Convert.ToDouble(textRowSplitted[3]),
                                        Convert.ToDouble(textRowSplitted[4]),
                                        Convert.ToDouble(textRowSplitted[5])
                                    );
                                    unstableJoints.Add(uniqueName, joint);
                                    joint.GetLabelAndStorey(sapModel, elevationToStoreyMap);
                                }
                                else { joint = unstableJoints[textRowSplitted[1]]; }

                                // Add unstable dimension to joint
                                joint.AddUnstableDimension(textRowSplitted[2]);
                            }
                        }
                    }
                }

                if (unstableJoints.Count == 0)
                {
                    sapModel.GroupDef.Delete(grpNm);
                    throw new Exception("No error joints found");
                }
                #endregion

                #region Add Joints to Group and Create Write Object
                // Create Group
                int ret = sapModel.GroupDef.Delete(grpNm);
                ret = sapModel.GroupDef.SetGroup(grpNm);

                // Init Write Lists
                List<string> labelNames = new List<string>();
                List<string> uniqueNames = new List<string>();
                List<double> x = new List<double>();
                List<double> y = new List<double>();
                List<double> z = new List<double>();
                List<string> storeyNames = new List<string>();
                List<string> unstableDimension = new List<string>();
                List<string> status = new List<string>();

                // Iterate through joints
                foreach (EtabsJoint joint in unstableJoints.Values)
                {
                    #region Group Joints
                    if (joint.uniqueName[0] == '~')
                    {
                        joint.status = "Internal Joint";
                    }
                    else
                    {
                        ret = sapModel.PointObj.SetGroupAssign(joint.uniqueName, grpNm);
                        if (ret == 0)
                        {
                            joint.status = $"Added to {grpNm}";
                        }
                        else
                        {
                            joint.status = $"Unable to add to {grpNm}";
                        }
                    }
                    #endregion

                    #region Create Write Object
                    labelNames.Add(joint.labelName);
                    uniqueNames.Add(joint.uniqueName);
                    x.Add(joint.x / 1000);
                    y.Add(joint.y / 1000);
                    z.Add(joint.z / 1000);
                    unstableDimension.Add(joint.GetAllUnstableDimensionsAsString());
                    status.Add(joint.status);
                    storeyNames.Add(joint.storeyName);
                    #endregion
                }
                #endregion

                #region Write to Excel
                WriteToExcelSelectionAsRow(0, 0, true,
                    labelNames.ToArray(),
                    uniqueNames.ToArray(),
                    x.ToArray(),
                    y.ToArray(),
                    z.ToArray(),
                    storeyNames.ToArray(),
                    unstableDimension.ToArray(),
                    status.ToArray()
                    );

                MessageBox.Show("Completed", "Completed");
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void groupAndImportLog_Click_OG(object sender, EventArgs e)
        {
            // Copy from old code, to refractor later
            try
            {
                #region Init ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion

                #region Create Lists
                List<string> errorJoints = new List<string>();
                List<double> coord1 = new List<double>();
                List<double> coord2 = new List<double>();
                List<double> coord3 = new List<double>();
                string grpNm = ".E.Error Joints";
                #endregion

                #region Analyse Files
                string logFilePath = getAnalysisFile(sapModel, ".LOG");
                int numLines = File.ReadLines(logFilePath).Count();

                using (StreamReader sr = new StreamReader(logFilePath))
                {
                    for (int i = 0; i < numLines; i++)
                    {
                        string line = sr.ReadLine();
                        if (line.Length > 6)
                        {
                            if (line.Substring(1, 5) == "Joint")
                            {
                                string[] row = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                if (!(errorJoints.Contains(row[1])))
                                {
                                    errorJoints.Add(row[1]);
                                    coord1.Add(Convert.ToDouble(row[3]));
                                    coord2.Add(Convert.ToDouble(row[4]));
                                    coord3.Add(Convert.ToDouble(row[5]));
                                }
                            }
                        }

                    }
                }

                if (errorJoints.Count == 0)
                {
                    sapModel.GroupDef.Delete(grpNm);
                    throw new Exception("No error joints found");
                }
                #endregion

                #region Add Joints to Group
                int ret = sapModel.GroupDef.Delete(grpNm);
                ret = sapModel.GroupDef.SetGroup(grpNm);
                List<string> grouped = new List<string>();
                int counter = 0;
                foreach (string joint in errorJoints)
                {
                    if (joint[0] == '~')
                    {
                        grouped.Add("Internal Joint");
                    }
                    else
                    {
                        ret = sapModel.PointObj.SetGroupAssign(joint, grpNm);
                        if (ret == 0)
                        {
                            counter++;
                            grouped.Add("Added");
                        }
                        else
                        {
                            grouped.Add("Failed to Add");
                        }
                    }
                }
                #endregion

                #region Write to Excel
                WriteToExcelSelectionAsRow(0, 0, true, errorJoints.ToArray(), coord1.ToArray(), coord2.ToArray(), coord3.ToArray(), grouped.ToArray());
                MessageBox.Show("Completed", "Completed");
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void groupAndImportWrn_Click(object sender, EventArgs e)
        {
            try
            {
                throw new Exception("Not Implemented");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void openLog_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                string logFilePath = getAnalysisFile(sapModel, ".LOG");
                Process.Start(logFilePath);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void openWRN_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                string logFilePath = getAnalysisFile(sapModel, ".WRN");
                Process.Start(logFilePath);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private void findSlantedWalls_Click(object sender, EventArgs e)
        {
            // Copy from old code, to refractor later
            try
            {
                #region Init ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion

                #region Get list of walls from ETABS
                int NumberNames = -1;
                string[] WallNames = null;
                ETABSv1.eAreaDesignOrientation[] DesignOrientation = null;
                int NumberBoundaryPts = -1;
                int[] PointDelimiter = null;
                string[] PointNames = null;
                double[] PointX = null;
                double[] PointY = null;
                double[] PointZ = null;

                int ret = sapModel.AreaObj.GetAllAreas(ref NumberNames, ref WallNames, ref DesignOrientation, ref NumberBoundaryPts, ref PointDelimiter, ref PointNames, ref PointX, ref PointY, ref PointZ);
                #endregion

                #region Initialise new error group
                string grpName = ".E.Slanted Walls"; // Set group name for error list
                ret = sapModel.GroupDef.SetGroup(grpName);
                ret = sapModel.GroupDef.Delete(grpName);
                ret = sapModel.GroupDef.SetGroup(grpName);
                int NumWalls = 0;
                int numFailedWalls = 0;
                #endregion

                #region Analyse Walls
                // For each wall, compare the location of the coordinates and check whether there is a matching pair
                for (int i = 0; i < NumberNames; i++)
                {
                    if (DesignOrientation[i].ToString() == "Wall")
                    {
                        NumWalls++;
                        // Find Number of Points to loop Through
                        int numPoints = 0;
                        if (i == 0)
                        {
                            numPoints = PointDelimiter[i] + 1;
                        }
                        else
                        {
                            numPoints = PointDelimiter[i] - PointDelimiter[i - 1];
                        }

                        // Isolate required Points
                        double[] localX = new double[numPoints];
                        double[] localY = new double[numPoints];
                        double[] localZ = new double[numPoints];
                        int index = PointDelimiter[i] - numPoints + 1;
                        Array.Copy(PointX, index, localX, 0, numPoints);
                        Array.Copy(PointY, index, localY, 0, numPoints);
                        Array.Copy(PointZ, index, localZ, 0, numPoints);

                        // Round the numbers to 3 decimal place
                        int dp = 4;
                        for (int j = 0; j < localX.Count(); j++)
                        {
                            localX[j] = Math.Round(localX[j], dp, MidpointRounding.AwayFromZero);
                            localY[j] = Math.Round(localY[j], dp, MidpointRounding.AwayFromZero);
                            localZ[j] = Math.Round(localZ[j], dp, MidpointRounding.AwayFromZero);
                        }

                        // Count number of distinct points
                        int distinctX = localX.Distinct().Count();
                        int distinctY = localY.Distinct().Count();
                        int distinctZ = localZ.Distinct().Count();

                        if (((distinctX > 2) || (distinctY > 2) || (distinctZ > 2)))
                        {
                            // Wall is slanted add to Group
                            ret = sapModel.AreaObj.SetGroupAssign(WallNames[i], grpName);
                            numFailedWalls++;
                        }
                    }
                }
                #endregion

                #region Report Status
                string message = "Number of walls checked = " + NumWalls.ToString() + "\nNumber of walls failed = " + numFailedWalls.ToString();
                if (numFailedWalls > 0) { message += $"\nCheck walls in group: {grpName}"; }
                else { ret = sapModel.GroupDef.Delete(grpName); }
                MessageBox.Show(message, "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                #endregion
            }

            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region Story Table
        class StoryTable
        {
            object[,] contents;

            Dictionary<string, int> storyElevationToIndex = new Dictionary<string, int>();
            public StoryTable(Range tableRange)
            {
                MapTable(tableRange);
            }
            public void MapTable(Range tableRange)
            {
                contents = GetContentsAsObject2DArray(tableRange);
                CheckForDoubles(tableRange);

                for (int rowNum = 0; rowNum < contents.GetLength(0); rowNum++)
                {
                    double elevationDouble;
                    try
                    {
                        elevationDouble = double.Parse(contents[rowNum, 1].ToString());
                    }
                    catch { throw new Exception($"Unable to parse \"{contents[rowNum, 1]}\" into number"); }


                    string elevationString = elevationDouble.ToString("#.####");
                    if (elevationString == "") { elevationString = "0"; }
                    if (storyElevationToIndex.ContainsKey(elevationString)) { throw new Exception($"Duplicate elevation \"{elevationString}\" found in story table"); }
                    storyElevationToIndex.Add(elevationString, rowNum);
                }
            }

            #region Check Doubles
            public void CheckForDoubles(Range sourceRange)
            {
                Range firstCell = sourceRange.Cells[1, 1];
                // Only check from 2nd column onwards
                for (int i = 0; i < contents.GetLength(0); i++)
                {
                    for (int j = 1; j < contents.GetLength(1); j++)
                    {
                        object cellValue = contents[i, j];

                        if (!(cellValue is double))
                        {
                            throw new Exception($"Error: Value '{cellValue}' in cell {firstCell.Offset[i, j].Address[false, false]} is not a number.");
                        }
                    }
                }
            }

            #endregion

            #region Get Values
            private int GetIndexFromElevation(double elevation)
            {
                string elevationString = elevation.ToString("#.####");
                if (elevationString == "") { elevationString = "0"; }
                if (!storyElevationToIndex.ContainsKey(elevationString)) { throw new Exception($"Story elevation \"{elevationString}\" not found in story table"); }
                return storyElevationToIndex[elevationString];
            }
            public double EffHeight(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 2];
            }
            public double MinWL(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 3];
            }

            public double MaxWL(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 4];
            }

            public double MinX(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 5];
            }

            public double MaxX(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 6];
            }

            public double MinY(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 7];
            }

            public double MaxY(double elevation)
            {
                int rowNum = GetIndexFromElevation(elevation);
                return (double)contents[rowNum, 8];
            }

            #endregion

            #region Wind Load Equations
            public Func<double, double> WindLoadInY(double elevation)
            {
                double x1 = MinX(elevation);
                double x2 = MaxX(elevation);
                double y1 = MinWL(elevation);
                double y2 = MaxWL(elevation);
                // y = mx + c
                double m = (y2 - y1) / (x2 - x1);
                double c = y1 - m * x1; // c = y - mx

                Func<double, double> windLoadEquation = x => (m * x + c);
                return windLoadEquation;
            }

            public Func<double, double> WindLoadInX(double elevation)
            {
                double x1 = MinY(elevation);
                double x2 = MaxY(elevation);
                double y1 = MinWL(elevation);
                double y2 = MaxWL(elevation);
                // y = mx + c
                double m = (y2 - y1) / (x2 - x1);
                double c = y1 - m * x1; // c = y - mx

                Func<double, double> windLoadEquation = x => (m * x + c);
                return windLoadEquation;
            }
            #endregion
        }


        #endregion

        #region Test Retrieve Table
        private void getPilingForces_Click(object sender, EventArgs e)
        {
            try
            {
                //throw new Exception("Not Implemented");
                #region Init ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                (_, Dictionary<double, string> elevationToStoreyMap) = GetEtabsStoreys(sapModel);
                #endregion

                #region Get Tables
                int numberTables = 0;
                string[] tableKey = null;
                string[] tableName = null;
                int[] importType = null;

                int ret = sapModel.DatabaseTables.GetAvailableTables(
                    ref numberTables,
                    ref tableKey,
                    ref tableName,
                    ref importType
                );
                if ( ret != 0 ) { throw new Exception("Error retrieving table list"); }
                #endregion

                #region Get Column Data Table
                string[] fieldKeyList = null;
                string[] fieldsKeysIncluded = null;
                string[] tableData = null;
                int tableVersion = 0;
                int numberRecords = 0;

                ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    "Design Forces - Columns",
                    ref fieldKeyList,
                    "All",
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData
                );
                if (ret != 0) { throw new Exception("Error retrieving Column data table list"); }
                #endregion

                #region Get Pier Data Table
                ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                    "Design Forces - Piers",
                    ref fieldKeyList,
                    "All",
                    ref tableVersion,
                    ref fieldsKeysIncluded,
                    ref numberRecords,
                    ref tableData
                );
                if (ret != 0) { throw new Exception("Error retrieving Pier data table list"); }
                #endregion

            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }


        #endregion
    }
}
