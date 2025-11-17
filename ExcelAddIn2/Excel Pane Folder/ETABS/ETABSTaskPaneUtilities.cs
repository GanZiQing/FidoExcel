using Autodesk.AutoCAD.DatabaseServices;
using ETABSv1;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;
using static ExcelAddIn2.EtabsFunctions;

namespace ExcelAddIn2
{
    public partial class ETABSTaskPane : UserControl
    {
        #region Initialise
        private void CreateAttributesForUtilities()
        {
            CustomAttribute thisAtt;
            AttributeTextBox thisAttBox;

            thisAtt = new CheckBoxAttribute("printFrameLabel_EtabsUtil", printFrameLabelCheck, false);
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new CheckBoxAttribute("printFrameCoord_EtabsUtil", printFrameCoordCheck, false);
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new CheckBoxAttribute("printFrameSection_EtabsUtil", printFrameSectionCheck, false);
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new CheckBoxAttribute("getColCheck_EtabsUtil", getColCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);
            thisAtt = new CheckBoxAttribute("getBeamCheck_EtabsUtil", getBeamCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);
            thisAtt = new CheckBoxAttribute("getBraceCheck_EtabsUtil", getBraceCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);
            thisAtt = new CheckBoxAttribute("getNullFrameCheck_EtabsUtil", getNullFrameCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);
            thisAtt = new CheckBoxAttribute("getOtherFrameCheck_EtabsUtil", getOtherFrameCheck, true);
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAttBox = new AttributeTextBox("frameUnOffsetColNum_EtabsUtil", dispFrameUnOffsetColNum, true);
            thisAttBox.SetDefaultValue("1");
            thisAttBox.type = "int";
            attributeDic.Add(thisAttBox.attName, thisAttBox);
        }
        #endregion

        #region Frame Tools
        private void getFrameUn_Click(object sender, EventArgs e)
        {
            try
            {
                #region Initalise ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion


                #region Get Frames And Details
                (_, string[] frameUns) = GetSelectedElementsByType(sapModel, EtabsObjectType.Frame);
                EtabsFrame[] frames = EtabsObjectTypeHelper.GetSpecificFrameType(frameUns, sapModel, GetTargetFrameTypes());

                for (int i = 0; i < frames.Length; i++)
                {
                    EtabsFrame frame = frames[i];
                    EtabsObjectTypeHelper.GetFrameDetails(sapModel, ref frame,
                        printFrameLabelCheck.Checked,
                        printFrameCoordCheck.Checked,
                        printFrameSectionCheck.Checked
                        );
                }
                #endregion

                #region Create Write Object
                #region Size
                int colCount = 1;
                if (printFrameLabelCheck.Checked) { colCount++; }
                if (printFrameSectionCheck.Checked) { colCount++; }
                if (printFrameCoordCheck.Checked) { colCount += 6; }

                object[,] finalWriteArray = new object[frames.Length, colCount];
                #endregion

                for (int rowNum = 0; rowNum < frames.Length; rowNum++)
                {
                    EtabsFrame frame = frames[rowNum];

                    finalWriteArray[rowNum, 0] = frame.uniqueName;
                    int colNum = 1;
                    if (printFrameLabelCheck.Checked) { finalWriteArray[rowNum, colNum] = frame.labelName; colNum++; }
                    if (printFrameSectionCheck.Checked) { finalWriteArray[rowNum, colNum] = frame.sectionName; colNum++; }
                    if (printFrameCoordCheck.Checked) 
                    { 
                        for (int coordNum = 0; coordNum < frame.coordinates.Length; coordNum++)
                        {
                            finalWriteArray[rowNum, colNum] = frame.coordinates[coordNum];
                            colNum++;
                        }
                    }
                }
                #endregion

                WriteObjectToExcelRange(null, 0, 0, true, finalWriteArray);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void setFrameUn_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Input
                Range selectedRange = GetSelectedExcelRange();
                string[] oldUn = GetContentsAsStringArray(selectedRange, false);
                CheckRangeSize(selectedRange,0,1, "Frame Unique Names");

                int offsetColNum = ((AttributeTextBox)attributeDic["frameUnOffsetColNum_EtabsUtil"]).GetIntFromTextBox();
                if (offsetColNum < 1) { throw new Exception("Offset Column Number must be greater than or equal to 1."); }

                Range newNameRange = selectedRange.Offset[0, offsetColNum];
                string[] newUn = GetContentsAsStringArray(newNameRange, false);
                #endregion

                #region Check Inputs
                List<string> emptyNewUn = new List<string>();
                for (int i = 0; i < oldUn.Length; i++)
                {
                    if (string.IsNullOrEmpty(oldUn[i])) { continue; }
                    if (string.IsNullOrEmpty(newUn[i])) { emptyNewUn.Add(oldUn[i]); }
                }
                if (emptyNewUn.Count > 0)
                {
                    string errorMessage = "The following Frame Unique Names have empty New Unique Names:\n";
                    foreach (string frameUn in emptyNewUn)
                    {
                        errorMessage += $"- {frameUn}\n";
                    }
                    throw new Exception(errorMessage);
                }
                #endregion

                #region Initalise ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion

                #region Change UN
                EtabsFrame[] allFrames = new EtabsFrame[oldUn.Length];
                string[] status = new string[oldUn.Length];
                bool errEncountered = false;
                for (int i = 0; i < oldUn.Length; i++)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(oldUn[i])) { continue; }
                        EtabsFrame frame = new EtabsFrame(oldUn[i], "");
                        frame.SetUniqueName(sapModel, newUn[i], true);
                        allFrames[i] = frame;
                    }
                    catch (Exception ex)
                    {
                        status[i] = ex.Message;
                        errEncountered = true;
                    }
                }
                #endregion

                #region Create Write If Error
                if (!errEncountered)
                {
                    MessageBox.Show("All Frame Unique Names were successfully changed.", "Success");
                    return;
                }
                else
                {
                    MessageBox.Show("Errors encountered, please check status", "Success");
                }
                    WriteToExcelRangeAsCol(newNameRange, 0, 1, true, status);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #endregion

        #region Area Tools
        private void getWallUNBut_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                (int numSel, string[] objName) = GetSelectedElementsByType(sapModel, EtabsObjectType.Area);
                WriteToExcelRangeAsCol(null, 0, 0, true, objName);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getWallPierBut_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);

                (int numSel, string[] objNames) = GetSelectedElementsByType(sapModel, EtabsObjectType.Area);
                string[] pierNames = new string[objNames.Length];
                for (int i = 0; i < objNames.Length; i++)
                {
                    string pierName = "";
                    sapModel.AreaObj.GetPier(objNames[i], ref pierName);
                    pierNames[i] = pierName;
                }
                WriteToExcelRangeAsCol(null, 0, 0, true, objNames, pierNames);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void setWallPierBut_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Info
                Range activeRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                CheckRangeSize(activeRange, 0, 2, "Assign Pier Labels");
                object[,] excelValues = GetContentsAsObject2DArray(activeRange);
                string[] wallUNs = new string[excelValues.GetLength(0)];
                string[] pierLabels = new string[excelValues.GetLength(0)];

                for (int rowNum = 0; rowNum < excelValues.GetLength(0); rowNum++)
                {
                    wallUNs[rowNum] = excelValues[rowNum, 0].ToString();
                    pierLabels[rowNum] = excelValues[rowNum, 1].ToString();
                }
                #endregion

                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);

                #region Get all Pier Labels
                int ret = 0;
                HashSet<string> allPierLabels = new HashSet<string>();
                {
                    int numNames = 0;
                    string[] pierLabelsETABS = new string[0];
                    ret = sapModel.PierLabel.GetNameList(ref numNames, ref pierLabelsETABS);
                    if (numNames == 0) { allPierLabels = pierLabelsETABS.ToHashSet(); }
                }
                #endregion

                #region Assign To ETABS
                bool allSuccess = true;
                string[] status = new string[wallUNs.Length];
                for (int rowNum = 0; rowNum < wallUNs.Length; rowNum++)
                {
                    try
                    {
                        // if pier does not exist add pier
                        string wallUN = wallUNs[rowNum];
                        string pierLabel = pierLabels[rowNum];
                        if (!allPierLabels.Contains(pierLabel))
                        {
                            ret = sapModel.PierLabel.SetPier(pierLabel);
                            if (ret == 0) { allPierLabels.Add(pierLabel); }
                        }

                        // Assign Pier to wall
                        ret = sapModel.AreaObj.SetPier(wallUN, pierLabel);
                        if (ret != 0) { status[rowNum] = $"Error encountered: Unknown"; allSuccess = false; }
                        else { status[rowNum] = $"Completed"; }
                    }
                    catch (Exception ex)
                    {
                        status[rowNum] = $"Error encountered: {ex.Message}";
                        allSuccess = false;
                    }
                }
                #endregion

                #region Write Status to ETABS
                if (!allSuccess)
                {
                    WriteToExcelRangeAsCol(null, 0, 2, false, status);
                    MessageBox.Show("One or more errors encountered, please see status column", "Warning");
                }
                else
                {
                    MessageBox.Show("Completed", "Completed");
                }
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        #endregion

        #region Helpers
        private HashSet<eFrameDesignOrientation> GetTargetFrameTypes()
        {
            HashSet<eFrameDesignOrientation> targetFrameTypes = new HashSet<eFrameDesignOrientation>();
            if (getColCheck.Checked)
            {
                targetFrameTypes.Add(eFrameDesignOrientation.Column);
            }
            if (getBeamCheck.Checked)
            {
                targetFrameTypes.Add(eFrameDesignOrientation.Beam);
            }
            if (getBraceCheck.Checked)
            {
                targetFrameTypes.Add(eFrameDesignOrientation.Brace);
            }
            if (getNullFrameCheck.Checked)
            {
                targetFrameTypes.Add(eFrameDesignOrientation.Null);
            }
            if (getOtherFrameCheck.Checked)
            {
                targetFrameTypes.Add(eFrameDesignOrientation.Other);
            }
            return targetFrameTypes; 
        }
        #endregion
    }
}
