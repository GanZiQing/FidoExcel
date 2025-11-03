using Autodesk.AutoCAD.DatabaseServices;
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
        }
        #endregion

        #region Get Base Reactions
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

                #region Check Input is Ok
                #region Check Group Names
                int ret = 1;
                // Probably should break these into their own functions at some point
                {
                    int numberNames = 0;
                    string[] myName = new string[0];
                    ret = sapModel.GroupDef.GetNameList(ref numberNames, ref myName);
                    HashSet<string> groupNameSet = new HashSet<string>(myName);
                    foreach (string groupName in groupNames)
                    {
                        if (!groupNameSet.Contains(groupName))
                        {
                            throw new Exception("Group Name " + groupName + " does not exist in the model.");
                        }
                    }
                }
                #endregion

                #region Check Combo Names
                // May have to shifted to after UN LC

                #endregion

                #endregion
                #endregion

                #region Get All Base Reaction from ETABS
                //(string[,] tableData, string[] fieldKeysIncluded) = GetEtabsTable2D(sapModel, "Joint Reactions");
                //Range writeRange = WriteToExcelRangeAsRow(null, 0, 0, true, fieldKeysIncluded);
                //FormatHeader(writeRange);
                //WriteObjectToExcelRange(null, 1, 0, true, tableData);

                (Dictionary<string, string[]> tableDataDic, string[] fieldKeysIncluded2, int numberRecords) = GetEtabsTableDic(sapModel, "Joint Reactions");

                #region Make Reference Dictionary


                #region Make Load Cases Unique ("UniqueLcName")
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
                            else
                            {
                                modifier = stepNum;
                            }
                            #endregion

                            if (modifier == null) { uniqueLC[entry] = LcBaseName; }
                            else { uniqueLC[entry] = LcBaseName + "-" + modifier; }
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

                        if (!basejointTracker.ContainsKey(uniqueName))
                        {
                            EtabsJoint newJoint = new EtabsJoint(uniqueName, labelName);
                            basejointTracker.Add(uniqueName, newJoint);
                        }

                        EtabsJoint joint = basejointTracker[uniqueName];
                        {
                            joint.AddShearReaction(loadCase, i,
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

                #region Check Combo Name
                {
                    int numberNames = 0;
                    string[] myName = new string[0];
                    ret = sapModel.RespCombo.GetNameList(ref numberNames, ref myName);

                    HashSet<string> comboNameSet = new HashSet<string>(myName);
                    foreach (string comboName in comboNames)
                    {
                        if (tableDataDic["UniqueLcName"].Contains(comboName))
                        {
                            continue; // Load case is ok
                        }
                        else
                        {
                            string msg;
                            if (comboNameSet.Contains(comboName))
                            {
                                // Load case exist but excludes modifier
                                msg = "Case/Combo Name " + comboName + " exist but has multiple entries.";
                            }
                            else
                            {
                                // Load case does not exist
                                msg = "Case/Combo Name " + comboName + " does not exist in the model.";
                            }

                            MessageBox.Show($"{msg}\nUse \"Get Active Load Combos\" function to get valid load case/combo names", "Error");
                            //WriteToExcelRangeAsCol(null, 0, 0, true, tableDataDic["UniqueLcName"].ToHashSet().ToArray());
                            return;
                        }
                    }
                }
                #endregion
                #endregion

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
                            if (!basejointTracker.ContainsKey(jointUN)) { continue;  } // Skip joints not in base reaction table
                            EtabsJoint joint = basejointTracker[jointUN];
                            double[] jointBaseReactions = joint.baseReactions[comboName];
                            for (int i = 0; i < 6; i++)
                            {
                                totalBaseReactions[i] += jointBaseReactions[i];
                            }
                        }
                        #endregion

                        mapComboToReactionsT.Add(comboName, totalBaseReactions);
                    }
                    mapGroupToComboReactions.Add(groupName, mapComboToReactionsT);
                }
                #endregion

                #region Create Write Object
                int numDof = 3;
                object[,] finalWriteArray = new object[groupNames.Length + 2, comboNames.Length * numDof + 1];
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

                //WriteObjectToExcelRange(null, 0, 0, true, headerArray);
                //WriteObjectToExcelRange(null, headerArray.GetLength(0), 1, true, writeArray);
                WriteObjectToExcelRange(null, 0, 0, true, finalWriteArray);

                // Write Contents


                //object[,] writeArray = ConcatArrays(List<object[,]> writeDataArrayList);

                //for (int rowNum = 0; rowNum < writeObj.GetLength(0); rowNum++)
                //{
                //    for (int colNum = 0; colNum < writeObj.GetLength(0); colNum++)
                //    {
                //        writeObj[]
                //    }
                //}
                #endregion

                #region Print to Excel

                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
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
            Dictionary<string, GeneralEtabsObject> colAndWallObj = EtabsObjectTypeHelper.GetVerticalElements(sapModel, groupName);
            Dictionary<string, GeneralEtabsObject> jointsDic = EtabsObjectTypeHelper.GetJointsFromElements(sapModel, colAndWallObj.Values, false);

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
        private void testGetObjButt_Click(object sender, EventArgs e)
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
                // Probably should break these into their own functions at some point
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
                #region Get Input from Excel
                ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).CheckRangeSize(0, 1);
                string[] comboNames = ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).GetContentsAsStringArray(true);
                #endregion
                
                InitializeETABS(out cOAPI etabsObject, out cSapModel sapModel, true);

                #region Get All Base Reaction from ETABS

                (Dictionary<string, string[]> tableDataDic, string[] fieldKeysIncluded2, int numberRecords) = GetEtabsTableDic(sapModel, "Joint Reactions");

                #region Make Reference Dictionary

                #region Make Load Cases Unique ("UniqueLcName")
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
                            else
                            {
                                modifier = stepNum;
                            }
                            #endregion

                            if (modifier == null) { uniqueLC[entry] = LcBaseName; }
                            else { uniqueLC[entry] = LcBaseName + "-" + modifier; }
                        }
                        tableDataDic.Add("UniqueLcName", uniqueLC);
                    }
                    else
                    {
                        tableDataDic.Add("UniqueLcName", tableDataDic["OutputCase"]);
                    }
                }
                #endregion

                WriteToExcelRangeAsCol(null, 0, 0, true, tableDataDic["UniqueLcName"].ToHashSet().ToArray());
                #endregion
                
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        private void getAllReactionsButt_Click(object sender, EventArgs e)
        {
            try
            {
                // Probably more efficient way to do this but I'm reusing code from above for now
                #region Get Input from Excel
                ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).CheckRangeSize(0, 1);
                string[] comboNames = ((RangeTextBox)attributeDic["loadComboRange_BaseShear"]).GetContentsAsStringArray(true);
                #endregion

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
        #endregion
    }
}

