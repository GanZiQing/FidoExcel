using Autodesk.AutoCAD.DatabaseServices;
using ETABSv1;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using PdfSharp.Snippets.Font;
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
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Autodesk.AutoCAD.DatabaseServices.Ole2Frame;
using static ExcelAddIn2.CommonUtilities;
using static ExcelAddIn2.EtabsFunctions;

namespace ExcelAddIn2
{
    public partial class ETABSTaskPane : UserControl
    {

        #region Initialise
        private void CreateAttributesForBaseShear()
        {
            AttributeTextBox thisTBAtt = new RangeTextBox("groupRange_BaseShear", dispGroupRange, setGroupRange, "column", false);
            attributeDic.Add(thisTBAtt.attName, thisTBAtt);

            thisTBAtt = new RangeTextBox("loadComboRange_BaseShear", dispLcRange, setLcRange, "column", false);
            attributeDic.Add(thisTBAtt.attName, thisTBAtt);

            CustomAttribute thisAtt = new ComboBoxAttribute("getForcesObj_BaseShear", dispGetForcesObj, "Column & Wall");
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new ComboBoxAttribute("tableFormat_BaseShear", dispTableFormat, "Append Bottom");
            attributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new CheckBoxAttribute("outputMoments_BaseShear", printMomentsCheck, false);
            attributeDic.Add(thisAtt.attName, thisAtt);
        }
        #endregion

        #region Get Base Reactions (Main)
        private void getBaseShearButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Initalise ETABS
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                #endregion

                #region Get Input from Excel
                ((RangeTextBox)attributeDic["groupRange_BaseShear"]).CheckRangeSize(0, 1);
                ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).CheckRangeSize(0, 1);

                string[] groupNames = ((RangeTextBox)attributeDic["groupRange_BaseShear"]).GetContentsAsStringArray(true);
                string[] comboNames = ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).GetContentsAsStringArray(true);

                CheckEtabsGroupExists(sapModel, groupNames);
                #endregion

                #region Get All Joint Reactions
                (Dictionary<string, EtabsJoint> basejointTracker,_) = GetAllJointReaction(comboNames, sapModel, true);
                #endregion

                #region Iterate
                Dictionary<string, object> mapGroupToComboReactions = new Dictionary<string, object>();

                foreach (string groupName in groupNames)
                {
                    Dictionary<string, double[]> mapComboToReactionsT = new Dictionary<string, double[]>();

                    foreach (string comboName in comboNames)
                    {
                        #region Get Base Joints
                        Dictionary<string, GeneralEtabsObject> targetJointsDic = GetJointsToRetriveForces(sapModel, groupName);
                        #endregion

                        #region Sum Base Shear Reactions
                        double[] totalBaseReactions = new double[6];

                        foreach (string jointUN in targetJointsDic.Keys)
                        {
                            if (!basejointTracker.ContainsKey(jointUN)) { continue; } // Skip joints not in base reaction table
                            EtabsJoint joint = basejointTracker[jointUN];
                            // Linear Superposition
                            double[] jointBaseReactions = joint.baseReactions[comboName];
                            for (int i = 0; i < 6; i++)
                            {
                                totalBaseReactions[i] += jointBaseReactions[i];
                            }
                            // Add moment
                            if (!printMomentsCheck.Checked) { continue; }
                            joint.GetMomentAboutPoint(sapModel, comboName, 0, 0);
                            totalBaseReactions[3] += jointBaseReactions[6];
                            totalBaseReactions[4] += jointBaseReactions[7];
                            totalBaseReactions[5] += jointBaseReactions[8];
                        }
                        #endregion

                        mapComboToReactionsT.Add(comboName, totalBaseReactions);
                    }
                    mapGroupToComboReactions.Add(groupName, mapComboToReactionsT);
                }
                #endregion

                #region Write to Excel
                if (dispTableFormat.Text == "Append Right")
                {
                    WriteBaseReaction_AppendRight(groupNames, comboNames, mapGroupToComboReactions);
                }
                else if (dispTableFormat.Text == "Append Bottom")
                {
                    WriteBaseReaction_AppendBottom(groupNames, comboNames, mapGroupToComboReactions);
                }
                else
                {
                    MessageBox.Show("Unknown table format selected, defaulting to Append Bottom", "Warning");
                    WriteBaseReaction_AppendBottom(groupNames, comboNames, mapGroupToComboReactions);
                }
                #endregion

            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }
        private (Dictionary<string, EtabsJoint> basejointTracker, HashSet<string> allUniqueLcNames) GetAllJointReaction(string[] comboNames, cSapModel sapModel, bool checkComboNames)
        {
            
            (Dictionary<string, string[]> tableDataDic, string[] fieldKeysIncluded2, int numberRecords) = GetEtabsTableDic(sapModel, "Joint Reactions");

            #region Make Load Cases Unique ("UniqueLcName")
            HashSet<string> allUniqueLcNames = new HashSet<string>();
            {
                if (tableDataDic.Keys.Contains("StepNumber"))
                {
                    string[] uniqueLC = new string[numberRecords];
                    for (int entry = 0; entry < numberRecords; entry++)
                    {
                        string LcBaseName = tableDataDic["OutputCase"][entry];

                        #region Create Modifier
                        string stepType;
                        {
                            if (tableDataDic.ContainsKey("StepType")) { stepType = tableDataDic["StepType"][entry]; }
                            else { stepType = null; }
                        }

                        string stepNum;
                        {
                            if (tableDataDic.ContainsKey("StepNumber")) { stepNum = tableDataDic["StepNumber"][entry]; }
                            else { stepNum = null; }
                        }

                        string modifier;
                        if (stepNum == null && stepType == null)
                        {
                            modifier = null;
                        }
                        else if (stepNum == null)
                        {
                            modifier = stepType;
                        }
                        else if (stepType == null)
                        {
                            modifier = stepNum;
                        }
                        else
                        {
                            modifier = stepType + stepNum;
                        }
                        #endregion

                        if (modifier == null) { uniqueLC[entry] = LcBaseName; }
                        else { uniqueLC[entry] = LcBaseName + "-" + modifier; }

                        #region Only Add to Valid LC Names if not Max/Min
                        if (!(stepType == "Max" || stepType == "Min"))
                        {
                            allUniqueLcNames.Add(uniqueLC[entry]);
                        }
                        #endregion
                    }
                    tableDataDic.Add("UniqueLcName", uniqueLC);
                }
                else
                {
                    tableDataDic.Add("UniqueLcName", tableDataDic["OutputCase"]);
                }
            }
            #endregion

            #region Make Reference Dictionary
            Dictionary<string, EtabsJoint> basejointTracker = new Dictionary<string, EtabsJoint>();
            {

                for (int i = 0; i < numberRecords; i++)
                {
                    // Create Joint Object
                    string uniqueName = tableDataDic["UniqueName"][i];
                    string labelName = tableDataDic["Label"][i];
                    string loadCase = tableDataDic["UniqueLcName"][i];
                    if (tableDataDic.ContainsKey("StepType"))
                    {
                        string stepType = tableDataDic["StepType"][i];
                        if (stepType == "Max" || stepType == "Min")
                        {
                            continue; // Skip max/min entries, pure summation not valid
                        }
                    }

                    if (!basejointTracker.ContainsKey(uniqueName))
                    {
                        EtabsJoint newJoint = new EtabsJoint(uniqueName, labelName);
                        basejointTracker.Add(uniqueName, newJoint);
                    }

                    EtabsJoint joint = basejointTracker[uniqueName];
                    {
                        joint.AddBaseReaction(loadCase, i,
                        double.Parse(tableDataDic["FX"][i]),
                        double.Parse(tableDataDic["FY"][i]),
                        double.Parse(tableDataDic["FZ"][i]),
                        double.Parse(tableDataDic["MX"][i]),
                        double.Parse(tableDataDic["MY"][i]),
                        double.Parse(tableDataDic["MZ"][i])
                        );
                    }
                }
            }
            #endregion

            if (checkComboNames) { CheckComboNames(comboNames, allUniqueLcNames, sapModel); }
            return (basejointTracker, allUniqueLcNames);
        }
        private void CheckComboNames(string[] comboNames, HashSet<string> allUniqueLcNames, cSapModel sapModel)
        {
            int ret = -1;
            int numberNames = 0;
            string[] myName = new string[0];
            ret = sapModel.RespCombo.GetNameList(ref numberNames, ref myName);

            HashSet<string> comboNameSet = new HashSet<string>(myName);
            foreach (string comboName in comboNames)
            {
                if (allUniqueLcNames.Contains(comboName))
                {
                    continue; // Load case is ok
                }
                else
                {
                    string msg;
                    if (comboNameSet.Contains(comboName))
                    {
                        // Load case exist but excludes modifier
                        msg = "Case/Combo Name " + comboName + " exist but is not valid for selection (e.g. envelope) or is not selected in ETABS table interface.";
                    }
                    else
                    {
                        // Load case does not exist
                        msg = "Case/Combo Name " + comboName + " does not exist in the model.";
                    }

                    throw new Exception($"{msg}\nUse \"Get Active Load Combos\" function to get valid load case/combo names\"");
                }
            }
        }

        private Range WriteBaseReaction_AppendBottom(string[] groupNames, string[]comboNames, Dictionary<string, object> mapGroupToComboReactions)
        {
            #region Find Write Object Size
            int numDof = printMomentsCheck.Checked ? 6 : 3;
            object[,] finalWriteArray = new object[groupNames.Length * comboNames.Length + 1, numDof + 2];
            #endregion

            #region Create Write Object

            #region Header
            {
                object[,] headerArray;
                {
                    headerArray = new object[1, finalWriteArray.GetLength(1)];
                    
                    headerArray[0, 0] = "Group Name";
                    headerArray[0, 1] = "Load Combo/Pattern";
                    headerArray[0, 2] = "FX";
                    headerArray[0, 3] = "FY";
                    headerArray[0, 4] = "FZ";
                    if (printMomentsCheck.Checked)
                    {
                        headerArray[0, 5] = "MX";
                        headerArray[0, 6] = "MY";
                        headerArray[0, 7] = "MZ";
                    }
                }
                TwoDArrayFunctions.WriteArrayIntoArray(ref finalWriteArray, headerArray, 0, 0);
            }
            #endregion

            #region Contents
            List<object[,]> groupReactions = new List<object[,]>();

            int rowNum = 1;
            foreach (string groupName in groupNames)
            {
                Dictionary<string, double[]> mapComboToReactions = (Dictionary<string, double[]>)mapGroupToComboReactions[groupName];
                
                
                foreach (string comboName in comboNames)
                {
                    // Write Group Name and Combo Name
                    finalWriteArray[rowNum, 0] = groupName;
                    finalWriteArray[rowNum, 1] = comboName;

                    int colNum = 2;
                    object[,] reactions = new object[1, numDof];
                    for (int reactionNum = 0; reactionNum < numDof; reactionNum++)
                    {
                        finalWriteArray[rowNum, reactionNum + colNum] = mapComboToReactions[comboName][reactionNum];
                    }
                    rowNum += 1;
                }
            }
            #endregion
            #endregion

            RemoveNaNFromArray(ref finalWriteArray, "#NaN");
            Range writeRange = WriteObjectToExcelRange(null, 0, 0, true, finalWriteArray);
            return writeRange;
        }

        private Range WriteBaseReaction_AppendRight(string[] groupNames, string[] comboNames, Dictionary<string, object> mapGroupToComboReactions)
        {
            #region Find Write Object Size
            int numDof = printMomentsCheck.Checked ? 6 : 3;
            object[,] finalWriteArray = new object[groupNames.Length + 2, comboNames.Length * numDof + 1];
            #endregion
            
            #region Create Write Object
            #region Header
            // Write Header Row
            {
                object[,] headerArray;
                {
                    headerArray = new object[2, finalWriteArray.GetLength(1)];
                    headerArray[0, 0] = "Load Case";
                    headerArray[1, 0] = "Group Name";
                    int colNum = 1;
                    foreach (string comboName in comboNames)
                    {
                        headerArray[0, colNum] = comboName;
                        headerArray[1, colNum] = "FX";
                        headerArray[1, colNum + 1] = "FY";
                        headerArray[1, colNum + 2] = "FZ";

                        colNum += numDof;
                    }
                }
                TwoDArrayFunctions.WriteArrayIntoArray(ref finalWriteArray, headerArray, 0, 0);
            }
            #endregion

            #region Group Names
            {
                object[,] loadCaseArray = new object[groupNames.Length, 1];
                int i = 0;
                foreach (string groupName in groupNames)
                {
                    loadCaseArray[i, 0] = groupName;
                    i += 1;
                }
                TwoDArrayFunctions.WriteArrayIntoArray(ref finalWriteArray, loadCaseArray, 2, 0);
            }
            #endregion

            #region Contents
            List<object[,]> groupReactions = new List<object[,]>();
            foreach (string groupName in groupNames)
            {
                Dictionary<string, double[]> mapComboToReactions = (Dictionary<string, double[]>)mapGroupToComboReactions[groupName];

                List<object[,]> comboReactions = new List<object[,]>();
                foreach (string comboName in comboNames)
                {
                    object[,] reactions = new object[1, numDof];
                    for (int i = 0; i < numDof; i++)
                    {
                        reactions[0, i] = mapComboToReactions[comboName][i];
                    }
                    comboReactions.Add(reactions);
                }
                object[,] comboRecationArray = TwoDArrayFunctions.ConcatArraysBeside(comboReactions);
                groupReactions.Add(comboRecationArray);
            }
            object[,] writeArray = TwoDArrayFunctions.ConcatArraysBelow(groupReactions);
            TwoDArrayFunctions.WriteArrayIntoArray(ref finalWriteArray, writeArray, 2, 1);
            #endregion
            #endregion

            RemoveNaNFromArray(ref finalWriteArray, "#NaN");
            Range writeRange = WriteObjectToExcelRange(null, 0, 0, true, finalWriteArray);
            return writeRange;
        }
        #endregion

        #region Get Joints To Analyse
        private Dictionary<string, GeneralEtabsObject> GetJointsToRetriveForces(cSapModel sapModel, string groupName)
        {
            bool getFromColWall = GetForceOrigin();
            if (getFromColWall)
            {
                return GetJointsFromGroupedWallCol(sapModel, groupName);
            }
            else
            {
                return GetJointsFromGroup(sapModel, groupName);
            }
        }

        private Dictionary<string, GeneralEtabsObject> GetJointsFromGroup(cSapModel sapModel, string groupName)
        {
            // Current implementation only gets joints from selected group, not wall object
            (int[] objectTypeIds, string[] objectNames) = GetGroupElementofType(sapModel, groupName, new HashSet<EtabsObjectType> { EtabsObjectType.Point });
            Dictionary<string, GeneralEtabsObject> jointsDic = new Dictionary<string, GeneralEtabsObject>();
            foreach (string objectName in objectNames)
            {
                if (!jointsDic.ContainsKey(objectName))
                {
                    jointsDic.Add(objectName, new EtabsJoint(objectName, ""));
                }
            }
            return jointsDic;
        }

        private Dictionary<string, GeneralEtabsObject> GetJointsFromGroupedWallCol(cSapModel sapModel, string groupName)
        {
            List<GeneralEtabsObject> colAndWallObj = EtabsObjectTypeHelper.GetVerticalElements(sapModel, groupName);
            Dictionary<string, GeneralEtabsObject> jointsDic = EtabsObjectTypeHelper.GetUniqueJointsFromElements(sapModel, colAndWallObj, false);

            return jointsDic;
        }

        /// <summary>
        /// Returns true if forces are to be obtained at the base joints of the column/wall objects
        /// Returns false if forces are to be obtained from joints
        /// </summary>
        /// <returns></returns>
        private bool GetForceOrigin()
        {
            if (((ComboBoxAttribute)attributeDic["getForcesObj_BaseShear"]).attValue == "Column & Wall") { return true; }
            else { return false; }
        }


        #endregion

        #region Checkers
        private void getJointsUsedButt_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Input from Excel
                ((RangeTextBox)attributeDic["groupRange_BaseShear"]).CheckRangeSize(0, 1);
                string[] groupNames = ((RangeTextBox)attributeDic["groupRange_BaseShear"]).GetContentsAsStringArray(true);
                #endregion

                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                CheckEtabsGroupExists(sapModel, groupNames);
                
                WriteWarning();
                WriteToExcelRangeAsRow(null, 0, 0, false, groupNames);
                int colNum = 0;
                foreach (string groupName in groupNames)
                {
                    Dictionary<string, GeneralEtabsObject> joints = GetJointsToRetriveForces(sapModel, groupName);

                    WriteToExcelRangeAsCol(null, 1, colNum, false, joints.Keys.ToArray());
                    colNum += 1;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getActiveLoadComboButt_Click(object sender, EventArgs e)
        {
            try
            {
                // Probably more efficient way to do this but I'm reusing code from above for now
                //#region Get Input from Excel
                //((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).CheckRangeSize(0, 1);
                //string[] comboNames = ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).GetContentsAsStringArray(true);
                //#endregion

                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);

                (_, HashSet<string> allUniqueLcNames) = GetAllJointReaction(new string[0], sapModel, false);
                string[] sortedLcNames = allUniqueLcNames.ToArray();
                Array.Sort(sortedLcNames);
                WriteToExcelRangeAsCol(null, 0, 0, true, sortedLcNames);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getAllReactionsButt_Click(object sender, EventArgs e)
        {
            try
            {
                // Probably more efficient way to do this but I'm reusing code from above for now
                //#region Get Input from Excel
                //((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).CheckRangeSize(0, 1);
                //string[] comboNames = ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).GetContentsAsStringArray(true);
                //#endregion

                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);

                #region Get All Base Reaction from ETABS
                (string[,] tableData, string[] fieldKeysIncluded) = GetEtabsTable2D(sapModel, "Joint Reactions");
                WriteWarning();
                Range writeRange = WriteToExcelRangeAsRow(null, 0, 0, false, fieldKeysIncluded);
                FormatHeader(writeRange);
                WriteObjectToExcelRange(null, 1, 0, false, tableData);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void checkObjectsAreUnique_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Input from Excel
                ((RangeTextBox)attributeDic["groupRange_BaseShear"]).CheckRangeSize(0, 1);
                string[] groupNames = ((RangeTextBox)attributeDic["groupRange_BaseShear"]).GetContentsAsStringArray(true);
                #endregion

                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);

                #region Check Group Names
                int ret = 1;
                {
                    int numberNames = 0;
                    string[] myName = new string[0];
                    ret = sapModel.GroupDef.GetNameList(ref numberNames, ref myName);
                    HashSet<string> groupNameSet = new HashSet<string>(myName);
                    foreach (string groupName in groupNames)
                    {
                        if (!groupNameSet.Contains(groupName))
                        {
                            throw new Exception($"Group Name \"{groupName}\" does not exist in the model.");
                        }
                    }
                }
                #endregion

                #region Get all objects and assign to groups
                Dictionary<string, List<string>> frameObjs = new Dictionary<string, List<string>>(); // Key: Object Unique Name, Value: List of Group Names
                Dictionary<string, List<string>> areaObjs = new Dictionary<string, List<string>>();
                if (groupNames.Length == 1)
                {
                    MessageBox.Show("Only 1 group selected, no check done", "Error");
                    return;
                }

                foreach (string groupName in groupNames)
                {
                    List<GeneralEtabsObject> colAndWalls = EtabsObjectTypeHelper.GetVerticalElements(sapModel, groupName);
                    foreach (GeneralEtabsObject obj in colAndWalls)
                    {
                        if (obj.objectType == EtabsObjectType.Area)
                        {
                            if (!frameObjs.ContainsKey(obj.uniqueName))
                            {
                                frameObjs.Add(obj.uniqueName, new List<string>());
                            }
                            frameObjs[obj.uniqueName].Add(groupName);
                        }
                        else if (obj.objectType == EtabsObjectType.Frame)
                        {
                            if (!areaObjs.ContainsKey(obj.uniqueName))
                            {
                                areaObjs.Add(obj.uniqueName, new List<string>());
                            }
                            areaObjs[obj.uniqueName].Add(groupName);
                        }
                        else
                        {
                            throw new Exception("Non-frame or area object found in checker");
                        }
                    }
                }
                #endregion

                #region Find duplicated objects
                List<string> duplicatedFrames = new List<string>();
                
                foreach (KeyValuePair<string, List<string>> kvp in frameObjs)
                {
                    if (kvp.Value.Count > 1)
                    {
                        duplicatedFrames.Add(kvp.Key);
                    }
                }

                List<string> duplicatedAreas = new List<string>();
                foreach (KeyValuePair<string, List<string>> kvp in areaObjs)
                {
                    if (kvp.Value.Count > 1)
                    {
                        duplicatedAreas.Add(kvp.Key);
                    }
                }
                #endregion

                #region Return if no error
                if (duplicatedFrames.Count == 0 && duplicatedAreas.Count == 0)
                {
                    MessageBox.Show("All objects are unique across the specified groups.", "Check Complete");
                    return;
                }
                #endregion

                #region Write object if error
                WriteWarning("Duplicate objects found.");

                #region Write Frames
                // Header
                WriteToExcelRangeAsRow(null, 0, 0, false, new string[1] {"Frame Objects"});
                WriteToExcelRangeAsRow(null, 1, 0, false, duplicatedFrames.ToArray());
                // Body
                int colNum = 0;
                foreach (string uniqueName in duplicatedFrames)
                {
                    List<string> duplicatedInGroup = frameObjs[uniqueName];
                    WriteToExcelRangeAsCol(null, 2, colNum, false, duplicatedInGroup.ToArray());
                    colNum += 1;
                }
                #endregion

                #region Write Areas
                // Header
                colNum += 1;
                WriteToExcelRangeAsRow(null, 0, colNum, false, new string[1] { "Shell Objects" });
                WriteToExcelRangeAsRow(null, 1, colNum, false, duplicatedAreas.ToArray());
                // Body
               
                foreach (string uniqueName in duplicatedAreas)
                {
                    List<string> duplicatedInGroup = areaObjs[uniqueName];
                    WriteToExcelRangeAsCol(null, 2, colNum, false, duplicatedInGroup.ToArray());
                    colNum += 1;
                }
                #endregion

                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getGroupNames_Click(object sender, EventArgs e)
        {
            try
            {
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);
                HashSet<string> groupNames = GetExistingEtabsGroup(sapModel, false);
                
                string[] printGroupNames = groupNames.ToArray();
                Array.Sort(printGroupNames);

                WriteToExcelRangeAsCol(null, 0, 0, true, printGroupNames);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #endregion
    }
}


