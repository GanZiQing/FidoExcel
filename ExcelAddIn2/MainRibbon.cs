﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
// Things I added myself
//using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using Office = Microsoft.Office.Core;

//using System.Threading;
// Things I added myself

namespace ExcelAddIn2
{
    public partial class MainRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

            Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            if (actCell.Value2 != null)
            {
                string sValue = actCell.Value2.ToString();
                string sText = actCell.Text;
                System.Windows.Forms.MessageBox.Show(sText);
            }
            //System.Threading.Thread.CurrentThread.CurrentCulture = original_Language;
        }

        private bool InitializeETABS(out ETABSv1.cOAPI etabsObject, out ETABSv1.cSapModel sapModel)
        {
            bool attachToInstance = true;
            etabsObject = null;
            sapModel = default(ETABSv1.cSapModel);

            if (attachToInstance)
            {
                // Attach to a running instance of ETABS 
                try
                {
                    // Get the active ETABS object
                    etabsObject = (ETABSv1.cOAPI)Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");
                }
                catch //(Exception ex)
                {
                    MessageBox.Show("No running instance of the program found or failed to attach.");
                    return false;
                }
            }

            // Get a reference to cSapModel to access all API classes and functions
            sapModel = etabsObject.SapModel;

            if (sapModel == null)
            {
                // Handle the case where SapModel is null
                return false;
            }

            return true;
        }

        private void SetPointUniqueName_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject; ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }

            // Main code starts from this line:

            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            Range oldNameRange = objSheet.Range[objSheet.Cells[rng.Row, rng.Column], objSheet.Cells[rng.Row + rng.Rows.Count - 1, rng.Column]];
            Range newNameRange = objSheet.Range[objSheet.Cells[rng.Row, rng.Column + 1], objSheet.Cells[rng.Row + rng.Rows.Count - 1, rng.Column + 1]];

            object[,] oldNames = oldNameRange.Value2;
            object[,] newNames = newNameRange.Value2;

            // Disable Excel screen updating
            objBook.Application.ScreenUpdating = false;

            // Change the unique names of selected points in ETABS
            int rowCount = rng.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                string oldName = oldNames[i, 1]?.ToString();
                string newName = newNames[i, 1]?.ToString();

                if (!string.IsNullOrEmpty(oldName) && !string.IsNullOrEmpty(newName))
                {
                    int ret = mySapModel.PointObj.ChangeName(oldName, newName);
                }
            }

            // Enable Excel screen updating
            objBook.Application.ScreenUpdating = true;

            // Clean up variables related to Excel
            objSheet = null;

            MessageBox.Show("Coding completed successfully", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void GetPointUniqueName_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }

            // Main code starts from this line:
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            string[] pointNames = new string[6869];
            int pointCount = 0;
            bool pointSelected = false;
            int ret = 0;

            // Get the names of selected points in ETABS
            ret = mySapModel.PointObj.GetNameList(ref pointCount, ref pointNames);

            // Execute the code without Task.Run
            Range Title = objSheet.Cells[rng.Row, rng.Column];
            Title.Value = "UniqueName";
            Range startCell = objSheet.Cells[rng.Row + 1, rng.Column];

            objBook.Application.ScreenUpdating = false;

            object[,] dataArray = new object[pointCount, 1];

            int pointIndex = 0;
            for (int i = 0; i < pointNames.Count(); i++)
            {
                ret = mySapModel.PointObj.GetSelected(pointNames[i], ref pointSelected);
                if (pointSelected)
                {
                    dataArray[pointIndex, 0] = pointNames[i];
                    pointIndex++;
                }
            }

            // Write the entire array to the worksheet in one go using Value2
            Range endCell = startCell.Offset[pointCount - 1, 0];
            Range writeRange = objSheet.Range[startCell, endCell];
            writeRange.Value2 = dataArray;

            objBook.Application.ScreenUpdating = true;

            objSheet = null;

            MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetGroups_Click(object sender, RibbonControlEventArgs e)
        {
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }

            // Main code starts here
            // Get group names from ETABS
            int ret = 0;
            int NumberNames = 0;
            string[] MyName = new string[0];
            ret = mySapModel.GroupDef.GetNameList(ref NumberNames, ref MyName);

            // Print to excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            Range Title = objSheet.Cells[rng.Row, rng.Column];
            Title.Value = "Group Names";
            Range startCell = objSheet.Cells[rng.Row + 1, rng.Column];

            objBook.Application.ScreenUpdating = false;

            object[,] dataArray = new object[NumberNames, 1];

            int pointIndex = 0;
            for (int i = 0; i < NumberNames; i++) // Come back to fix this, I think we don't need if loop
            {
                dataArray[pointIndex, 0] = MyName[i];
                pointIndex++;
            }

            // Write the entire array to the worksheet in one go using Value2
            Range endCell = startCell.Offset[NumberNames - 1, 0];
            Range writeRange = objSheet.Range[startCell, endCell];
            writeRange.Value2 = dataArray;
            //objSheet.Cells[rng.Row, rng.Column + 1].Value = "To Replicate"; This is how you write to one cell only 

            objBook.Application.ScreenUpdating = true;

            objSheet = null;

            MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetPointCoord_Click(object sender, RibbonControlEventArgs e)
        {
            // Obselete
            //// Common code to initiate the Etabs
            //ETABSv1.cOAPI myETABSObject;
            //ETABSv1.cSapModel mySapModel;

            //if (!InitializeETABS(out myETABSObject, out mySapModel))
            //{
            //    // Handle initialization failure
            //    return;
            //}

            //// Main code starts here

            //// Add section to read input data from Excel
            //Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            //Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            //// Get points from ETABS
            //int ret = 0;
            //int pointCount = 0;
            //string[] pointNames = new string[1];
            //bool pointSelected = false;
            //ret = mySapModel.PointObj.GetNameList(ref pointCount, ref pointNames);

            //string[] selectedJoints = new string[pointCount];
            //double[] x = new double[pointCount];
            //double[] y = new double[pointCount];
            //double[] z = new double[pointCount];
            //int selPointCount = 0;
            //for (int i = 0; i < pointCount; i++)
            //{
            //    ret = mySapModel.PointObj.GetSelected(pointNames[i], ref pointSelected);
            //    if (pointSelected)
            //    {
            //        selectedJoints[selPointCount] = pointNames[i];
            //        ret = mySapModel.PointObj.GetCoordCartesian(pointNames[i], ref x[selPointCount], ref y[selPointCount], ref z[selPointCount]);
            //        selPointCount++;
            //    }
            //}

            //// Write to Excel
            //int numCol = 4;
            //int numRow = selPointCount;

            //objBook.Application.ScreenUpdating = false;
            //object[,] dataArray = new object[selPointCount, 4];


            //for (int j = 0; j < numCol; j++) // Columns
            //{
            //    for (int i = 0; i < numRow; i++) // Rows
            //    {
            //        if (j == 0)
            //        {
            //            dataArray[i, j] = pointNames[i];
            //        }
            //        else if (j == 1)
            //        {
            //            dataArray[i, j] = x[i];
            //        }
            //        else if (j == 2)
            //        {
            //            dataArray[i, j] = y[i];
            //        }
            //        else if (j == 3)
            //        {
            //            dataArray[i, j] = z[i];
            //        }
            //    }
            //}

            //// Write the entire array to the worksheet in one go using Value2
            //Range startCell = objSheet.Cells[rng.Row, rng.Column];
            //Range endCell = startCell.Offset[numRow-1, numCol-1];
            //Range writeRange = objSheet.Range[startCell, endCell];
            //writeRange.Value2 = dataArray;

            //objBook.Application.ScreenUpdating = true;
            //objSheet = null;

            //MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void getNodeCoord_Click(object sender, RibbonControlEventArgs e)
        {
            // Print to excel the joint unique name and x,y,z of the joints selected in ETABS
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Main code starts here
            // Add section to read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            // Get points from ETABS
            int ret = 0;
            int NumSel = 0;
            int[] ObjectType = new int[0];
            string[] ObjectName = new string[0];
            ret = mySapModel.SelectObj.GetSelected(ref NumSel, ref ObjectType, ref ObjectName);

            int selPointCount = 0;
            double[] x = new double[NumSel];
            double[] y = new double[NumSel];
            double[] z = new double[NumSel];
            string[] selectedJoints = new string[NumSel];

            for (int i = 0; i < NumSel; i++)
            {
                if (ObjectType[i] == 1)
                {
                    selectedJoints[selPointCount] = ObjectName[i];
                    ret = mySapModel.PointObj.GetCoordCartesian(ObjectName[i], ref x[selPointCount], ref y[selPointCount], ref z[selPointCount]);
                    selPointCount++;
                }
            }
            // Truncate results for printing
            string[] selectedJoints2 = new string[selPointCount];
            double[] x2 = new double[selPointCount];
            double[] y2 = new double[selPointCount];
            double[] z2 = new double[selPointCount];

            Array.Copy(selectedJoints, 0, selectedJoints2, 0, selPointCount);
            Array.Copy(x, 0, x2, 0, selPointCount);
            Array.Copy(y, 0, y2, 0, selPointCount);
            Array.Copy(z, 0, z2, 0, selPointCount);
            
            // Print
            WriteToExcel(0,0,selectedJoints2,x2,y2,z2);
            MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DuplicateUnits_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }

            // Main code starts here
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);
            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            // User inputs
            bool ReadAsFixed = false; // set to false to read from selection 
            bool addToLatestGroup = true;

            // Reading data from excel
            int lastRow = 0;
            int lastCol = 0;
            int firstRow = 0;
            int firstCol = 0;
            Range dataRange = objSheet.Range[objSheet.Cells[1, 1], objSheet.Cells[1, 1]];
            if (ReadAsFixed)
            {
                lastRow = objSheet.Cells[objSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
                lastCol = 10;
                firstRow = 3;
                firstCol = 1;
                dataRange = objSheet.Range[objSheet.Cells[firstRow, firstCol], objSheet.Cells[lastRow, lastCol]];
            }
            else
            {
                dataRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
                lastRow = dataRange.Rows.Count;
                lastCol = dataRange.Columns.Count;
                firstRow = 1;
                firstCol = 1;
            }

            // Check for correct number of columns
            if (!(((lastCol == 8) | (lastCol == 10)) | (lastCol == 12)))
            {
                MessageBox.Show("Wrong number of columns, only 8, 10 or 12 allowed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Read Excel data as object
            object[,] data = dataRange.Value2;

            // Convert to individual arrays
            string[] unitLabel = new string[lastRow - firstRow + 1];
            string[] unitGroup = new string[lastRow - firstRow + 1];
            double[] refX = new double[lastRow - firstRow + 1];
            double[] refY = new double[lastRow - firstRow + 1];
            double[] refZ = new double[lastRow - firstRow + 1];
            double[] targX = new double[lastRow - firstRow + 1];
            double[] targY = new double[lastRow - firstRow + 1];
            double[] targZ = new double[lastRow - firstRow + 1];
            double[] rot = new double[lastRow - firstRow + 1];
            string[] mirr = new string[lastRow - firstRow + 1];
            bool[] groupAsUnit = new bool[lastRow - firstRow + 1];
            bool[] groupAsElement = new bool[lastRow - firstRow + 1];

            for (int i = 1; i < lastCol + 1; i++)
            {
                switch (i)
                {
                    case 1:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            unitLabel[j - 1] = data[j, i]?.ToString();
                        }
                        break;

                    case 2:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            unitGroup[j - 1] = data[j, i]?.ToString();
                        }
                        break;

                    case 3:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            refX[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;

                    case 4:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            refY[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;

                    case 5:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            targX[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;

                    case 6:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            targY[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;

                    case 7:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            rot[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;

                    case 8:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            mirr[j - 1] = Convert.ToString(data[j, i]);
                        }
                        break;

                    case 9:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            if (Convert.ToString(data[j, i]) == "Yes")
                            {
                                groupAsUnit[j - 1] = true;
                            }
                            else
                            {
                                groupAsUnit[j - 1] = false;
                            }
                        }
                        break;

                    case 10:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            if (Convert.ToString(data[j, i]) == "No")
                            {
                                groupAsElement[j - 1] = false;
                            }
                            else
                            {
                                groupAsElement[j - 1] = true;
                            }
                        }
                        break;

                    case 11:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            refZ[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;

                    case 12:
                        for (int j = 1; j <= lastRow - firstRow + 1; j++)
                        {
                            targZ[j - 1] = Convert.ToDouble(data[j, i]);
                        }
                        break;
                }
            }
            // Calculate dZ
            double[] dZ = new double[lastRow - firstRow + 1];
            for (int i = 0; i < targZ.Count(); i++)
            {
                dZ[i] = targZ[i] - refZ[i];
            }

            // Deal with ETABS
            int ret = 0;

            // Create reset group for latest group
            string latestGrpName = "";
            if (addToLatestGroup)
            {
                latestGrpName = "..Last Duplicated";
                ret = mySapModel.GroupDef.Delete(latestGrpName);
                ret = mySapModel.GroupDef.SetGroup(latestGrpName);
            }

            // Loop through all rows of data in excel in ETABS
            for (int i = 0; i < lastRow - firstRow + 1; i++)
            {
                // Copy elements:
                if (lastCol == 8)
                {
                    CopyElements(mySapModel, unitGroup[i], unitLabel[i], targX[i], targY[i], refX[i], refY[i], rot[i], mirr[i]);
                }
                else if (lastCol == 10)
                {
                    CopyElements(mySapModel, unitGroup[i], unitLabel[i], targX[i], targY[i], refX[i], refY[i], rot[i], mirr[i], groupAsUnit[i], groupAsElement[i], latestGrpName);
                }
                else if (lastCol == 12)
                {
                    CopyElements(mySapModel, unitGroup[i], unitLabel[i], targX[i], targY[i], refX[i], refY[i], rot[i], mirr[i], groupAsUnit[i], groupAsElement[i], latestGrpName, dZ[i]);
                }
            }
            ret = mySapModel.View.RefreshView();
            MessageBox.Show("Coding completed successfully", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string[] GetStoryNames(ETABSv1.cSapModel mySapModel)
        {
            double BaseElevation = 0;
            int NumberStories = 0;
            string[] StoryNames = new string[0];
            double[] StoryElevations = new double[0];
            double[] StoryHeights = new double[0];
            bool[] IsMasterStory = new bool[0];
            string[] SimilarToStory = new string[0];
            bool[] SpliceAbove = new bool[0];
            double[] SpliceHeight = new double[0];
            int[] color = new int[0];

            int ret = -1;
            ret = mySapModel.Story.GetStories_2(ref BaseElevation, ref NumberStories, ref StoryNames, ref StoryElevations, ref StoryHeights, ref IsMasterStory, ref SimilarToStory, ref SpliceAbove, ref SpliceHeight, ref color);
            return StoryNames;
        }
        private void GetJointLoad_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }

            // Main code starts here

            // Get Load from ETABS
            int ret = 0;
            int NumberItems = -1;
            string[] PointName = new string[0];
            string[] LoadPat = new string[0];
            int[] LCStep = new int[0];
            string[] CSys = new string[0];
            double[] F1 = new double[0];
            double[] F2 = new double[0];
            double[] F3 = new double[0];
            double[] M1 = new double[0];
            double[] M2 = new double[0];
            double[] M3 = new double[0];
            ETABSv1.eItemType ItemType = ETABSv1.eItemType.Group;

            ret = mySapModel.PointObj.GetLoadForce("MyGroup", ref NumberItems, ref PointName, ref LoadPat, ref LCStep, ref CSys, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3, ItemType);

            MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetFloors_Click(object sender, RibbonControlEventArgs e)
        {
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel SapModel;

            if (!InitializeETABS(out myETABSObject, out SapModel))
            {
                // Handle initialization failure
                return;
            }

            // Main code starts here
            // Get group names from ETABS
            int ret = 0;
            double BaseElevation = 0;
            int NumberStories = 0;
            string[] StoryNames = new string[0];
            double[] StoryElevations = new double[0];
            double[] StoryHeights = new double[0];
            bool[] IsMasterStory = new bool[0];
            string[] SimilarToStory = new string[0];
            bool[] SpliceAbove = new bool[0];
            double[] SpliceHeight = new double[0];
            int[] color = new int[0];

            ret = SapModel.Story.GetStories_2(ref BaseElevation, ref NumberStories, ref StoryNames, ref StoryElevations, ref StoryHeights, ref IsMasterStory, ref SimilarToStory, ref SpliceAbove, ref SpliceHeight, ref color);

            // Print to excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            objBook.Application.ScreenUpdating = false;

            // User Inputs
            int excelNoCol = 3;
            int startRowOffset = 1;
            string[] headers = { "Story Names", "Elevations", "Height" };

            // Write Title Blocks
            for (int i = 0; i < excelNoCol; i++)
            {
                objSheet.Cells[rng.Row, rng.Column + i].Value = headers[i];
                objSheet.Cells[rng.Row, rng.Column + i].Font.Bold = true;
                objSheet.Cells[rng.Row, rng.Column + i].Interior.Color = 16247773;
            }

            // Create Object with desired data
            object[,] dataArray = new object[NumberStories, excelNoCol];
            for (int i = 0; i < NumberStories; i++)
            {
                dataArray[i, 0] = StoryNames[i];
                dataArray[i, 1] = StoryElevations[i];
                dataArray[i, 2] = StoryHeights[i];
            }

            // Write the entire array to the worksheet in one go using Value2
            Range startCell = objSheet.Cells[rng.Row + startRowOffset, rng.Column];
            Range endCell = startCell.Offset[NumberStories - 1, excelNoCol - 1]; // -1 because it's an offset
            Range writeRange = objSheet.Range[startCell, endCell];
            writeRange.Value2 = dataArray;

            objBook.Application.ScreenUpdating = true;
            objSheet = null;

            MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CopyFrameLabel_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;

            // Get user input for groups
            var result = MessageBox.Show("Copy all groups?", "User Input", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            bool copyGroups = true;
            switch (result)
            {
                case DialogResult.No
                :
                    copyGroups = false;
                    break;
            }

            // Add section to read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            // Read Excel data as object
            Range dataRange = objSheet.Range[objSheet.Cells[rng.Row, rng.Column], objSheet.Cells[rng.Row + rng.Rows.Count, rng.Column + rng.Columns.Count]];
            object[,] data = dataRange.Value2;

            // Convert data to individual arrays
            //string[] StoryNames = GetStoryNames(mySapModel); this code is to get all stories
            string[] StoryNames = new string[rng.Rows.Count];
            for (int j = 1; j < rng.Columns.Count + 1; j++) // j refers to the column number in Excel
            {
                switch (j)
                {
                    case 1: // reading column 1
                        for (int i = 1; i < rng.Rows.Count + 1; i++)
                        {
                            StoryNames[i - 1] = data[i, j]?.ToString();
                        }
                        break;
                }
            }

            // Get the names of selected frames in ETABS
            int NumberNames = 0;
            string[] allFrameNames = new string[0];

            ret = mySapModel.FrameObj.GetNameList(ref NumberNames, ref allFrameNames);

            bool isFrameSelected = false;
            List<string> selectedFrames = new List<string>();

            for (int i = 0; i < NumberNames; i++)
            {
                ret = mySapModel.FrameObj.GetSelected(allFrameNames[i], ref isFrameSelected);
                if (isFrameSelected)
                {
                    selectedFrames.Add(allFrameNames[i]);
                }
            }

            // Get story data - This section is out dated
            //string[] StoryNames = GetStoryNames(mySapModel);

            // Duplicate the frame unique names to all other frame names
            foreach (string frameName in selectedFrames)
            {
                // Get global label
                string Label = "";
                string originalStory = "";
                ret = mySapModel.FrameObj.GetLabelFromName(frameName, ref Label, ref originalStory);

                // Get Global Group Assignment
                int NumberGroups = 0;
                string[] Groups = new string[0];
                ret = mySapModel.FrameObj.GetGroupAssign(frameName, ref NumberGroups, ref Groups);

                // Duplicate to each unique frame in that global label
                foreach (string story in StoryNames)
                {
                    string uniqueStringName = "";
                    ret = mySapModel.FrameObj.GetNameFromLabel(Label, story, ref uniqueStringName);
                    if (uniqueStringName != "")
                    {
                        string newUniqueName = frameName + "." + story;
                        ret = mySapModel.FrameObj.ChangeName(uniqueStringName, newUniqueName);

                        // Add to group
                        if (copyGroups)
                        {
                            for (int i = 0; i < NumberGroups; i++)
                            {
                                ret = mySapModel.FrameObj.SetGroupAssign(newUniqueName, Groups[i]);
                            }
                        }
                    }
                }
            }
            ret = mySapModel.View.RefreshView();
            MessageBox.Show("Coding completed successfully 20231128", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RemoveUNBack_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;

            // Get the names of selected frames in ETABS
            int NumberNames = 0;
            string[] allFrameNames = new string[0];

            ret = mySapModel.FrameObj.GetNameList(ref NumberNames, ref allFrameNames);

            bool isFrameSelected = false;
            List<string> selectedFrames = new List<string>();

            for (int i = 0; i < NumberNames; i++)
            {
                ret = mySapModel.FrameObj.GetSelected(allFrameNames[i], ref isFrameSelected);
                if (isFrameSelected)
                {
                    selectedFrames.Add(allFrameNames[i]);
                }
            }

            // Duplicate the frame unique names to all other frame names
            bool chkDuplicateName = false; // checker for whether duplicate frame name already exists
            bool chkNoBack = false; // checker for whether some names do not have the "." character 
            foreach (string frameName in selectedFrames)
            {
                int indexOfChar = frameName.LastIndexOf(".");
                if (indexOfChar != -1)
                {
                    string newUniqueName = frameName.Substring(0, indexOfChar);
                    // Check if frame name exist
                    string frameLabel = "";
                    string Story = "";
                    ret = mySapModel.FrameObj.GetLabelFromName(newUniqueName, ref frameLabel, ref Story);

                    if (ret != 0)
                    {
                        // The frame exists, change name
                        ret = mySapModel.FrameObj.ChangeName(frameName, newUniqueName);
                    }
                    else
                    {
                        chkDuplicateName = true;
                    }
                }
                else
                {
                    chkNoBack = true;
                }
            }
            ret = mySapModel.View.RefreshView();

            if (chkDuplicateName)
            {
                MessageBox.Show("Some Frames not renamed due to duplicate name", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (chkNoBack)
            {
                MessageBox.Show("Some Frames not duplicated as substring cannot be found ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            MessageBox.Show("Coding completed successfully", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SelectBeamLabel_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;

            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            bool checkRange = !((selectedRange.Rows.Count == 1) & (selectedRange.Columns.Count == 1));
            if (checkRange)
            {
                MessageBox.Show("Select one cell only", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            // Get story data
            string[] StoryNames = GetStoryNames(mySapModel);

            // Duplicate the frame unique names to all other frame names
            string selectedBeam = selectedRange.Value2;
            foreach (string story in StoryNames)
            {
                string uniqueName = "";
                ret = mySapModel.FrameObj.GetNameFromLabel(selectedBeam, story, ref uniqueName);
                ret = mySapModel.FrameObj.SetSelected(uniqueName, true);
                uniqueName = null;
            }
            //ret = mySapModel.View.RefreshView();
        }

        private void checkWalls_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;

            // Get list of walls from ETABS
            int NumberNames = -1;
            string[] WallNames = null;
            ETABSv1.eAreaDesignOrientation[] DesignOrientation = null;
            int NumberBoundaryPts = -1;
            int[] PointDelimiter = null;
            string[] PointNames = null;
            double[] PointX = null;
            double[] PointY = null;
            double[] PointZ = null;

            ret = mySapModel.AreaObj.GetAllAreas(ref NumberNames, ref WallNames, ref DesignOrientation, ref NumberBoundaryPts, ref PointDelimiter, ref PointNames, ref PointX, ref PointY, ref PointZ);

            // Initialise new error group
            string GrpName = ".E.Slanted Walls"; // Set group name for error list
            ret = mySapModel.GroupDef.SetGroup(GrpName);
            ret = mySapModel.GroupDef.Delete(GrpName);
            ret = mySapModel.GroupDef.SetGroup(GrpName);
            int NumWalls = 0;
            int NumFailedWalls = 0;
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
                        ret = mySapModel.AreaObj.SetGroupAssign(WallNames[i], GrpName);
                        NumFailedWalls++;
                    }
                }
            }
            string Contents = "Number of walls checked = " + NumWalls.ToString() + "\nNumber of walls failed = " + NumFailedWalls.ToString();
            MessageBox.Show(Contents, "Code Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void drawDropPanel_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            // Reading data from excel
            Range dataRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            int lastRow = dataRange.Rows.Count;
            int lastCol = dataRange.Columns.Count;
            int firstRow = 1;
            int firstCol = 1;

            // Check for correct number of columns
            if (!(lastCol == 9))
            {
                MessageBox.Show("Wrong number of columns, only 9 allowed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Read Excel data as object
            object[,] data = dataRange.Value2;

            // Convert to individual arrays
            double[] X = new double[lastRow - firstRow + 1];
            double[] Y = new double[lastRow - firstRow + 1];
            double[] Z = new double[lastRow - firstRow + 1];
            double[] dpX = new double[lastRow - firstRow + 1];
            double[] dpY = new double[lastRow - firstRow + 1];
            string[] sectionNm = new string[lastRow - firstRow + 1];
            double[] rotation = new double[lastRow - firstRow + 1];

            for (int i = firstCol; i < lastCol + 1; i++)
            {
                for (int j = 1; j <= lastRow - firstRow + 1; j++)
                {
                    switch (i)
                    {
                        case 2:
                            X[j - 1] = Math.Round(Convert.ToDouble(data[j, i]), 4);
                            break;

                        case 3:
                            Y[j - 1] = Math.Round(Convert.ToDouble(data[j, i]), 4);
                            break;

                        case 4:
                            Z[j - 1] = Math.Round(Convert.ToDouble(data[j, i]), 4);
                            break;

                        case 6:
                            dpX[j - 1] = Math.Round(Convert.ToDouble(data[j, i]), 4);
                            break;

                        case 7:
                            dpY[j - 1] = Math.Round(Convert.ToDouble(data[j, i]), 4);
                            break;

                        case 8:
                            sectionNm[j - 1] = Convert.ToString(data[j, i]);
                            break;
                        case 9:
                            rotation[j - 1] = Math.Round(Convert.ToDouble(data[j, i]), 4);
                            break;
                    }

                }
            }

            // Create reset group for latest group
            string latestGrpName = "..Last Duplicated";
            ret = mySapModel.GroupDef.Delete(latestGrpName);
            ret = mySapModel.GroupDef.SetGroup(latestGrpName);

            // Add DP at all nodes
            for (int i = 0; i < X.Count(); i++)
            {
                // Generate Basic Points
                int NumberPoints = 4;
                double[] xList_S = new double[NumberPoints]; // these are the input coordinates for the new shell
                double[] yList_S = new double[NumberPoints];
                double[] zList_S = new double[NumberPoints];
                for (int j = 0; j < NumberPoints; j++)
                {
                    if (j == 0)
                    {
                        xList_S[j] = X[i] - dpX[i] / 2;
                        yList_S[j] = Y[i] - dpY[i] / 2;
                        zList_S[j] = Z[i];
                    }
                    else if (j == 1)
                    {
                        xList_S[j] = X[i] + dpX[i] / 2;
                        yList_S[j] = Y[i] - dpY[i] / 2;
                        zList_S[j] = Z[i];
                    }
                    else if (j == 2)
                    {
                        xList_S[j] = X[i] + dpX[i] / 2;
                        yList_S[j] = Y[i] + dpY[i] / 2;
                        zList_S[j] = Z[i];
                    }
                    else if (j == 3)
                    {
                        xList_S[j] = X[i] - dpX[i] / 2;
                        yList_S[j] = Y[i] + dpY[i] / 2;
                        zList_S[j] = Z[i];
                    }
                }
                // Rotate points
                if (rotation[i] != 0)
                {
                    for (int j = 0; j < NumberPoints; j++)
                    {
                        (xList_S[j], yList_S[j], zList_S[j]) = CalculateNewCoordinates(xList_S[j], yList_S[j], zList_S[j], X[i], Y[i], 0, X[i], Y[i], rotation[i], "No");
                    }
                }

                // Add area
                string finalName_S = "";
                string PropName_S = sectionNm[i];
                ret = mySapModel.AreaObj.AddByCoord(NumberPoints, ref xList_S, ref yList_S, ref zList_S, ref finalName_S, PropName_S);
                ret = mySapModel.AreaObj.SetGroupAssign(finalName_S, latestGrpName);

            }
            //ret = mySapModel.View.RefreshView();

            // Completed
            MessageBox.Show("Coding completed successfully", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void errorJoints_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Read warning file for model
            string filePath = "";
            filePath = mySapModel.GetModelFilename();
            string errorFileName = filePath.Substring(0, filePath.Length - 4)+".LOG";
            int numLines = File.ReadLines(errorFileName).Count(); 
            string[] contents = new string[numLines];

            try
            {
                using (StreamReader sr = new StreamReader(errorFileName))
                {
                    for (int i = 0; i < numLines; i++)
                    {
                        string line = sr.ReadLine();
                        contents[i] = line;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            //// Only required if we want to read the other parts, but we are only reading up to the coordinates
            //// Define the table index splitter based on the " -----" lines first
            //string [] tableDelim = new string[0]; // list contianing the index of each
            //int [] tableIndex = new int[0];
            //foreach (string line in contents)
            //{
            //    if (line.Length > 4)
            //    {
            //        if (line.Substring(1, 3) == "---")
            //        {
            //            tableDelim = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            //            tableIndex = new int[tableDelim.Length];
            //            int col = 0;
            //            for (int index = 1; index < line.Length; index++)
            //            {
                          
            //                if ((line[index] == '-') & (line[index-1] == ' '))
            //                {
            //                    tableIndex[col] = index;
            //                    col++;

            //                }
            //            }
            //            break;
            //        }
            //    }
            //}
            // Add joints to dictionary
            //Dictionary<string, (double, double, double)> errorJoints = new Dictionary<string, (double, double, double)>();
            List<string> errorJoints = new List<string>();
            List<double> coord1 = new List<double>();
            List<double> coord2 = new List<double>();
            List<double> coord3 = new List<double>();
            foreach (string line in contents)
            {
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
            string GrpNm = "..Error Joints";
            ret = mySapModel.GroupDef.Delete(GrpNm);
            ret = mySapModel.GroupDef.SetGroup(GrpNm);
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
                    ret = mySapModel.PointObj.SetGroupAssign(joint, GrpNm);
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

            WriteToExcel(0,0,errorJoints.ToArray(), coord1.ToArray(), coord2.ToArray(), coord3.ToArray(), grouped.ToArray());
            string msgText = "Coding completed, " + counter.ToString() + " added.";
            MessageBox.Show(msgText, "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void offsetter_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range dataRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            int lastRow = dataRange.Rows.Count;
            int lastCol = dataRange.Columns.Count;
            int firstRow = 1;
            int firstCol = 1;

            // Read Excel data as object
            object[,] data = dataRange.Value2;

            // Convert to individual arrays
            string[] offsetType = new string[lastCol - firstCol + 1]; 
            bool[] addSlab = new bool[lastCol - firstCol + 1];
            double[] X = new double[lastCol - firstCol + 1];
            double[] Y = new double[lastCol - firstCol + 1];
            string[,] joints = new string[lastRow - firstRow + 1-4, lastCol - firstCol + 1];
            int[] numJoints = new int[lastCol - firstCol + 1];

            for (int colNum = 0; colNum < lastCol; colNum++) 
            {
                int eColNum = colNum + 1;
                for (int rowNum = 0; rowNum < lastRow; rowNum++)
                {
                    int eRowNum = rowNum + 1;
                    switch (eRowNum)
                    {
                        case 1:
                            offsetType[colNum] = Convert.ToString(data[eRowNum, eColNum]);
                            break;
                        case 2:
                            if (Convert.ToString(data[eRowNum, eColNum]) == "Yes")
                            {
                                addSlab[colNum] = true;
                            }
                            else
                            {
                                addSlab[colNum] = false;
                            }
                            break;
                        case 3:
                            X[colNum] = Convert.ToDouble(data[eRowNum, eColNum]);
                            break;
                        case 4:
                            Y[colNum] = Convert.ToDouble(data[eRowNum, eColNum]);
                            break;
                        default:
                            joints[rowNum-4, colNum] = Convert.ToString(data[eRowNum, eColNum]);
                            if (Convert.ToString(data[eRowNum, eColNum]) != null)
                            {
                                numJoints[colNum]++;
                            }
                            break;
                    }
                }
            }

            // Create reset group for latest group
            string latestGrpName = "..Last Duplicated";
            ret = mySapModel.GroupDef.Delete(latestGrpName);
            ret = mySapModel.GroupDef.SetGroup(latestGrpName);

            // Add DP at all nodes
            for (int colNum = 0; colNum < X.Count(); colNum++)
            {
                if (offsetType[colNum] == "1")
                {
                    // Add 4 points at each node
                    for (int rowNum = 0; rowNum < numJoints[colNum]; rowNum++) //Loop through each joint in that column
                    {
                        string joint = joints[rowNum, colNum]; // Origin joint 
                        double x = 0;
                        double y = 0;
                        double z = 0;
                        ret = mySapModel.PointObj.GetCoordCartesian(joint, ref x, ref y, ref z);
                        for (int countAddJoint = 0; countAddJoint < 4; countAddJoint++) // Loop through each direction
                        {
                            string jointName = "";
                            switch (countAddJoint)
                            {
                                case 0:
                                    ret = mySapModel.PointObj.AddCartesian(x - X[colNum], y + Y[colNum], z, ref jointName);
                                    break;
                                case 1:
                                    ret = mySapModel.PointObj.AddCartesian(x + X[colNum], y + Y[colNum], z, ref jointName);
                                    break;
                                case 2:
                                    ret = mySapModel.PointObj.AddCartesian(x - X[colNum], y - Y[colNum], z, ref jointName);
                                    break;
                                case 3:
                                    ret = mySapModel.PointObj.AddCartesian(x + X[colNum], y - Y[colNum], z, ref jointName);
                                    break;
                            }
                            if (ret == 0)
                            {
                                ret = mySapModel.PointObj.SetGroupAssign(jointName, latestGrpName);
                            }
                        }
                    }
                }
                else if (offsetType[colNum] == "4")
                {
                    // Get joint coordinates
                    double[] x = new double[4];
                    double[] y = new double[4];
                    double[] z = new double[4];
                    for (int rowNum = 0; rowNum < 4; rowNum++) //Loop through each joint in that column
                    {
                        ret = mySapModel.PointObj.GetCoordCartesian(joints[rowNum, colNum], ref x[rowNum], ref y[rowNum], ref z[rowNum]);
                    }

                    // Identify type of corner point
                    int[] cornerType = new int[4];
                    double minX = x.Min();
                    double minY = y.Min();
                    double maxX = x.Max();
                    double maxY = y.Max();
                    for (int rowNum = 0; rowNum < 4; rowNum++) //Loop through each joint in that column
                    {
                        if ((x[rowNum] == minX) & (y[rowNum] == maxY))
                        {
                            cornerType[rowNum] = 0;
                        }
                        else if ((x[rowNum] == maxX) & (y[rowNum] == maxY))
                        {
                            cornerType[rowNum] = 1;
                        }
                        else if ((x[rowNum] == minX) & (y[rowNum] == minY))
                        {
                            cornerType[rowNum] = 2;
                        }
                        else if ((x[rowNum] == maxX) & (y[rowNum] == minY))
                        {
                            cornerType[rowNum] = 3;
                        }
                        else
                        {
                            cornerType[rowNum] = -1;
                        }
                    }

                    // Duplicate one point for each joint based on type 
                    for (int rowNum = 0; rowNum < 4; rowNum++) //Loop through each joint in that column
                    {
                        string jointName = joints[rowNum, colNum];
                        switch (cornerType[rowNum])
                        {
                            case 0:
                                ret = mySapModel.PointObj.AddCartesian(x[rowNum] - X[colNum], y[rowNum] + Y[colNum], z[rowNum], ref jointName);
                                break;
                            case 1:
                                ret = mySapModel.PointObj.AddCartesian(x[rowNum] + X[colNum], y[rowNum] + Y[colNum], z[rowNum], ref jointName);
                                break;
                            case 2:
                                ret = mySapModel.PointObj.AddCartesian(x[rowNum] - X[colNum], y[rowNum] - Y[colNum], z[rowNum], ref jointName);
                                break;
                            case 3:
                                ret = mySapModel.PointObj.AddCartesian(x[rowNum] + X[colNum], y[rowNum] - Y[colNum], z[rowNum], ref jointName);
                                break;
                        }
                        if (ret == 0)
                        {
                            ret = mySapModel.PointObj.SetGroupAssign(jointName, latestGrpName);
                        }
                    }

                }
            }
                ret = mySapModel.View.RefreshView();
            // Completed
            MessageBox.Show("Coding completed successfully", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetJointFromFrame_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range dataRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            int lastRow = dataRange.Rows.Count;
            int lastCol = dataRange.Columns.Count;
            int firstRow = 0;
            int firstCol = 0;

            // Read Excel data as object
            object[,] data = dataRange.Value2;

            // Convert to individual arrays
            string[] frameLabels = new string[lastRow - firstRow];

            for (int colNum = firstCol; colNum < lastCol; colNum++)
            {
                int eColNum = colNum + 1;
                for (int rowNum = 0; rowNum < lastRow; rowNum++)
                {
                    int eRowNum = rowNum + 1;
                    switch (eColNum)
                    {
                        case 1:
                            frameLabels[rowNum] = Convert.ToString(data[eRowNum, eColNum]);
                            break;
                    }
                }
            }

            // Get Point Names
            string[] Point1 = new string[frameLabels.Count()];
            string[] Point2 = new string[frameLabels.Count()];

            for (int i = 0; i < frameLabels.Count(); i++)
            {
                ret = mySapModel.FrameObj.GetPoints(frameLabels[i], ref Point1[i], ref Point2[i]);
            }

            // Write to Excel
            // Find number of rows and columns
            int numRow = frameLabels.Count();
            int numCol = 2;

            // Initiate object
            object[,] dataArray = new object[numRow, numCol];
            for (int col = 0; col < numCol; col++)
            {
                for (int row = 0; row < numRow; row++)
                {
                    if (col == 0)
                    {
                        dataArray[row, col] = Point1[row];
                    }
                    else if (col == 1)
                    {
                        dataArray[row, col] = Point2[row];
                    }
                    
                }
            }
            
            // Write to Excel
            objBook.Application.ScreenUpdating = false;

            // Write the entire array to the worksheet in one go using Value2
            Range startCell = objSheet.Cells[dataRange.Row, dataRange.Column+1];
            Range endCell = startCell.Offset[numRow - 1, numCol - 1];
            Range writeRange = objSheet.Range[startCell, endCell];
            writeRange.Value2 = dataArray;

            objBook.Application.ScreenUpdating = true;
            objSheet = null;
        }

        private void setJCoord_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range dataRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            int lastRow = dataRange.Rows.Count;
            int lastCol = dataRange.Columns.Count;
            int firstRow = 0;
            int firstCol = 0;

            // Read Excel data as object
            object[,] data = dataRange.Value2;

            // Convert to individual arrays
            string[] oldJointUN = new string[lastRow - firstRow];
            string[] newJointUN = new string[lastRow - firstRow];

            for (int colNum = firstCol; colNum < lastCol; colNum++)
            {
                int eColNum = colNum + 1;
                for (int rowNum = 0; rowNum < lastRow; rowNum++)
                {
                    int eRowNum = rowNum + 1;
                    switch (eColNum)
                    {
                        case 1:
                            oldJointUN[rowNum] = Convert.ToString(data[eRowNum, eColNum]);
                            break;
                        case 2:
                            newJointUN[rowNum] = Convert.ToString(data[eRowNum, eColNum]);
                            break;
                    }
                }
            }
            // Set Joint UN
            int counter = 0;
            for (int i = 0; i < oldJointUN.Count(); i++)
            {
                ret = mySapModel.PointObj.ChangeName(oldJointUN[i], newJointUN[i]);
                if (ret == 0)
                {
                    counter++;
                }
            }

            // Refresh View
            ret = mySapModel.View.RefreshWindow();
            // Message Box
            string msgText = "Coding completed, " + oldJointUN.Count() + " joints found. " + counter.ToString() + " successfully added.";
            MessageBox.Show(msgText, "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetPierForces_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = objBook.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            int firstRow = selectedRange.Row;
            int firstCol = selectedRange.Column;
            int numRow = selectedRange.Rows.Count;
            int numCol = selectedRange.Columns.Count;
            int lastRow = firstRow + numRow - 1;
            int lastCol = firstCol + numCol - 1;

            int numOffset = 1;
            Range dataRange = objSheet.Range[objSheet.Cells[firstRow,firstCol], objSheet.Cells[lastRow, lastCol + numOffset]];

            // Read Excel data as object
            object[,] data = dataRange.Value2;

            // Convert to individual arrays + dictionary 
            string[] pierLabels = new string[lastRow - firstRow + 1]; // List will be used for printing purposes
            string[] loadCases = new string[lastRow - firstRow + 1];
            int num_output = 13; // Number of columns of data to extract 

            // We will primarily be using dictionary for data extraction 
            Dictionary<string, Dictionary<string, double[]>> DPierLC = new Dictionary<string, Dictionary<string, double[]>>(); // Outer dictionary of PierLabel:LoadCase
            // DPierLC is a dictionary with key = pier labels, value = dictionary "LCArray" 
            // LCArray is a dictionary with key = load case names, value = Array of size num_output containing loading values:
            //0 P max, 1 Pmin, 
            //2 M2b_min, 3 M2b_max, 4 M2t_min, 5 M2t_max, 
            //6 M3b_min, 7 M3b_max, 8 M3t_min, 9 M3t_max, 
            //10 V2_Max, 11 V3_Max, 12 T_Max

            for (int rowNum = 0; rowNum < numRow; rowNum++)
            {
                int eRowNum = rowNum + 1; // This is the row num in excel
                // Read data first
                pierLabels[rowNum] = Convert.ToString(data[eRowNum, 1]);
                loadCases[rowNum] = Convert.ToString(data[eRowNum, 2]);

                if (DPierLC.ContainsKey(pierLabels[rowNum]))
                {
                    //If exist, add new load case to key 
                    if (!DPierLC[pierLabels[rowNum]].ContainsKey(loadCases[rowNum]))
                    {
                        double[] LCArray = new double[num_output];
                        for (int i = 0; i < num_output; i++)
                        {
                            LCArray[i] = double.NaN;
                        }
                        DPierLC[pierLabels[rowNum]].Add(loadCases[rowNum], LCArray);
                    }
                }
                else
                {
                    //Create new key then 
                    Dictionary<string, double[]> DLCL = new Dictionary<string, double[]>();
                    double[] LCArray = new double[num_output];
                    for (int i = 0; i < num_output; i++)
                    {
                        LCArray[i] = double.NaN;
                    }
                    DLCL.Add(loadCases[rowNum], LCArray);
                    DPierLC.Add(pierLabels[rowNum], DLCL);
                }
            }
            // Enable all LC
            EnableAllLC(mySapModel);
            // Technically you can only enable the loadcases found in the output you require

            // Get ETABS result for all piers
            int NumberResults = 0;
            string[] StoryName = null;
            string[] PierName = null;
            string[] LoadCase = null;
            string[] Location = null;
            double[] P = null;
            double[] V2 = null;
            double[] V3 = null;
            double[] T = null;
            double[] M2 = null;
            double[] M3 = null;
            ret = mySapModel.Results.PierForce(ref NumberResults, ref StoryName, ref PierName, ref LoadCase, ref Location, ref P, ref V2, ref V3, ref T, ref M2, ref M3);

            // Iterate through all output, save result if Pier and LC exist within dictionary
            for (int i=0; i < NumberResults; i++)
            {
                if (DPierLC.ContainsKey(PierName[i]))
                {
                    if (DPierLC[PierName[i]].ContainsKey(LoadCase[i]))
                    {
                        /*--------------------PMax--------------------*/
                        if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][0]) || P[i] < DPierLC[PierName[i]][LoadCase[i]][0])
                        {
                            DPierLC[PierName[i]][LoadCase[i]][0] = P[i];
                        }
                        /*--------------------PMin--------------------*/
                        if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][1]) || P[i] > DPierLC[PierName[i]][LoadCase[i]][1]) //Reversed for P because P downwards is negative
                        {
                            DPierLC[PierName[i]][LoadCase[i]][1] = P[i];
                        }
                        /*--------------------M2b&t--------------------*/
                        if (Location[i] == "Bottom")
                        {
                            if (M2[i] < 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][2]) || M2[i] < DPierLC[PierName[i]][LoadCase[i]][2])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][2] = M2[i]; //M2b_Min
                                }
                            }
                            else if (M2[i] > 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][3]) || M2[i] > DPierLC[PierName[i]][LoadCase[i]][3])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][3] = M2[i]; //M2b_Max
                                }
                            }
                        }
                        else if (Location[i] == "Top")
                        {
                            if (M2[i] < 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][4]) || M2[i] < DPierLC[PierName[i]][LoadCase[i]][4])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][4] = M2[i]; //M2t_Min
                                }
                            }
                            else if (M2[i] > 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][5]) || M2[i] > DPierLC[PierName[i]][LoadCase[i]][5])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][5] = M2[i]; //M2t_Max
                                }
                            }
                        }
                        /*--------------------M3b&t--------------------*/
                        if (Location[i] == "Bottom")
                        {
                            if (M3[i] < 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][6]) || M3[i] < DPierLC[PierName[i]][LoadCase[i]][6])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][6] = M3[i]; //M3b_Min
                                }
                            }
                            else if (M3[i] > 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][7]) || M3[i] > DPierLC[PierName[i]][LoadCase[i]][7])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][7] = M3[i]; //M3b_Max
                                }
                            }
                        }
                        else if (Location[i] == "Top")
                        {
                            if (M3[i] < 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][8]) || M3[i] < DPierLC[PierName[i]][LoadCase[i]][8])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][8] = M3[i]; //M3t_Min
                                }
                            }
                            else if (M3[i] > 0)
                            {
                                if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][9]) || M3[i] > DPierLC[PierName[i]][LoadCase[i]][9])
                                {
                                    DPierLC[PierName[i]][LoadCase[i]][9] = M3[i]; //M3t_Max
                                }
                            }
                        }
                        /*--------------------V2 V3 T--------------------*/
                        if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][10]) || Math.Abs(V2[i]) > DPierLC[PierName[i]][LoadCase[i]][10])
                        {
                            DPierLC[PierName[i]][LoadCase[i]][10] = Math.Abs(V2[i]); //V2
                        }

                        if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][11]) || Math.Abs(V3[i]) > DPierLC[PierName[i]][LoadCase[i]][11])
                        {
                            DPierLC[PierName[i]][LoadCase[i]][11] = Math.Abs(V3[i]);
                        }

                        if (Double.IsNaN(DPierLC[PierName[i]][LoadCase[i]][12]) || Math.Abs(T[i]) > DPierLC[PierName[i]][LoadCase[i]][12])
                        {
                            DPierLC[PierName[i]][LoadCase[i]][12] = Math.Abs(T[i]);
                        }
                    }
                }
            }
            // Convert dict to arrays + Post processing
            double[] PMax = new double[pierLabels.Count()];
            double[] PMin = new double[pierLabels.Count()];
            double[] M2b = new double[pierLabels.Count()];
            double[] M2t = new double[pierLabels.Count()];
            double[] M3b = new double[pierLabels.Count()];
            double[] M3t = new double[pierLabels.Count()];
            double[] V2_max = new double[pierLabels.Count()];
            double[] V3_max = new double[pierLabels.Count()];
            double[] T_max = new double[pierLabels.Count()];

            double Clean_Zeros(double num) // Helper function
            {
                if (double.IsNaN(num) || Math.Abs(num) <= 0.01)
                { return 0; }
                else { return num; }
            }
            double Get_Max_Abs(double num1, double num2) // Helper function
            {
                num1 = Clean_Zeros(num1);
                num2 = Clean_Zeros(num2);
                if (Math.Abs(num1) > Math.Abs(num2))
                { return num1; }
                else { return num2; }
            }

            for (int i=0; i<pierLabels.Count(); i++)
            {
                //For dictionary:
                //0 P max, 1 Pmin, 
                //2 M2b_min, 3 M2b_max, 4 M2t_min, 5 M2t_max, 
                //6 M3b_min, 7 M3b_max, 8 M3t_min, 9 M3t_max, 
                //10 V2_Max, 11 V3_Max, 12 T_Max

                //For excel
                //0 P max, 1 Pmin, 
                //2 M2b, 3 M2t, 4 M3b, 5 M3t, 
                //6 V2_Max, 7 V3_Max, 8 T_Max
                for (int j = 0; j < 9; j++)
                {
                    switch (j)
                    {
                        case 0: // Pmax
                            PMax[i] = Clean_Zeros(DPierLC[pierLabels[i]][loadCases[i]][0]);
                            break;
                        case 1: // Pmin
                            PMin[i] = Clean_Zeros(DPierLC[pierLabels[i]][loadCases[i]][1]);
                            break;
                        case 2: // M2b
                            M2b[i] = Get_Max_Abs(DPierLC[pierLabels[i]][loadCases[i]][2], DPierLC[pierLabels[i]][loadCases[i]][3]);
                            break;
                        case 3: // M2t
                            M2t[i] = Get_Max_Abs(DPierLC[pierLabels[i]][loadCases[i]][4], DPierLC[pierLabels[i]][loadCases[i]][5]);
                            break;
                        case 4: // M3b
                            M3b[i] = Get_Max_Abs(DPierLC[pierLabels[i]][loadCases[i]][6], DPierLC[pierLabels[i]][loadCases[i]][7]);
                            break;
                        case 5: // M3t
                            M3t[i] = Get_Max_Abs(DPierLC[pierLabels[i]][loadCases[i]][8], DPierLC[pierLabels[i]][loadCases[i]][9]);
                            break;
                        case 6: //V2_Max
                            V2_max[i] = Clean_Zeros(DPierLC[pierLabels[i]][loadCases[i]][10]);
                            break;
                        case 7: //V3_Max
                            V3_max[i] = Clean_Zeros(DPierLC[pierLabels[i]][loadCases[i]][11]);
                            break;
                        case 8: //T_Max
                            T_max[i] = Clean_Zeros(DPierLC[pierLabels[i]][loadCases[i]][12]);
                            break;
                    }
                }
            }
            
            // Write to excel
            WriteToExcel(0, 2, PMax, PMin, M2b, M2t, M3b, M3t, V2_max, V3_max,T_max);
            //WriteToExcel(0, 13, StoryName, PierName, LoadCase, Location, P, M2, M3 , V2, V3, T);
        }

        private void SetOuputLC_Click(object sender, RibbonControlEventArgs e)
        {
            // Common code to initiate the Etabs
            ETABSv1.cOAPI myETABSObject;
            ETABSv1.cSapModel mySapModel;

            if (!InitializeETABS(out myETABSObject, out mySapModel))
            {
                // Handle initialization failure
                return;
            }
            int ret = 0;
            mySapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);

            //// Get all load case
            int NumberNames = 0;
            string[] MyName = null;
            ret = mySapModel.LoadCases.GetNameList(ref NumberNames, ref MyName);
            // Enable all load case
            for (int i = 0; i < NumberNames; i++)
            {
                ret = mySapModel.Results.Setup.SetCaseSelectedForOutput(MyName[i]);
                bool sel = false;
                ret = mySapModel.Results.Setup.GetCaseSelectedForOutput(MyName[i], ref sel);
            }


            //// Get all load combi
            int NumberNames2 = 0;
            string[] MyName2 = null;
            ret = mySapModel.RespCombo.GetNameList(ref NumberNames2, ref MyName2);
            for (int i = 0; i < NumberNames; i++)
            {
                ret = mySapModel.Results.Setup.SetComboSelectedForOutput(MyName2[i]);
                bool sel2 = false;
                ret = mySapModel.Results.Setup.GetCaseSelectedForOutput(MyName2[i], ref sel2);
            }

            // Get ETABS result for all piers
            int NumberResults = 0;
            string[] StoryName = null;
            string[] PierName = null;
            string[] LoadCase = null;
            string[] Location = null;
            double[] P = null;
            double[] V2 = null;
            double[] V3 = null;
            double[] T = null;
            double[] M2 = null;
            double[] M3 = null;
            ret = mySapModel.Results.PierForce(ref NumberResults, ref StoryName, ref PierName, ref LoadCase, ref Location, ref P, ref V2, ref V3, ref T, ref M2, ref M3);
        }

        private void GetFiles_Click(object sender, RibbonControlEventArgs e)
        {
            // Read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = objBook.ActiveSheet;
            Worksheet sourceSheet = objBook.Sheets["Macro Input"];
            Range selectedRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;

            // Specify the directory path you want to loop through
            //string directoryPath = @"D:\Documents\TPY - 248002-01\05_ST Reports\ST1\ST01 Screenshot\Overall";
            string directoryPath = sourceSheet.Range["B1"].Value2;


            // Code to read from property
            string RangeofGet = "Property that you read";
            Range activeRange = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Range[RangeofGet];
            // Function to get all file paths
            void ListFilesAndFolders(string directory, List<string> in_directories, List<string> in_files)
            {
                // Get all directories and files within the specified directory
                string[] subDirectories = Directory.GetDirectories(directory);
                string[] fileList = Directory.GetFiles(directory);

                // Add directories and files to the lists
                in_directories.AddRange(subDirectories);
                in_files.AddRange(fileList);
                // Recursively call this method for each subdirectory
                foreach (string subDir in subDirectories)
                {
                    ListFilesAndFolders(subDir, in_directories, in_files);
                }
            }


            // Call the method to list files and folders
            List<string> directories = new List<string>();
            List<string> files = new List<string>();
            ListFilesAndFolders(directoryPath, directories, files);

            // Print files array
            string[] folder_name = new string[files.Count()];
            string[] file_name = new string[files.Count()];
            string[] full_path = new string[files.Count()];
            int i = 0;
            foreach (string file in files)
            {
                full_path[i] = file;
                file_name[i] = Path.GetFileName(file);
                folder_name[i] = Path.GetFileName(Path.GetDirectoryName(file));
                i++;
            }
            WriteToExcel(0, 0, full_path, folder_name, file_name);
        }











        #region Helper Functions
        private void WriteToExcel(int rowOff, int colOff,params Array[] arrays)
        {
            // This code takes any number of arrays (of various types) and outputs them into excel 
            // Output order depends on order of the input array
            // Output location is the first cell of the current selection, offset by rowOff and colOff

            // Find number of rows and columns
            int numRow = 0;
            int numCol = arrays.Length;
            for (int col = 0; col < arrays.Length; col++)
            {
                if (arrays[col].Length > numRow)
                {
                    numRow = arrays[col].Length; // Finds max number of rows out of all the various arrays
                }
            }

            // Initiate object
            object[,] dataArray = new object[numRow, numCol];
            for (int col = 0; col < arrays.Length; col++)
            {
                for (int row = 0; row < arrays[col].Length; row++)
                {
                    dataArray[row, col] = arrays[col].GetValue(row);
                }
            }

            // Add section to read input data from Excel
            Workbook objBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet objSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Range rng = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            
            // Write to Excel
            objBook.Application.ScreenUpdating = false;

            // Write the entire array to the worksheet in one go using Value2
            Range startCell = objSheet.Cells[rng.Row + rowOff, rng.Column + colOff];
            Range endCell = startCell.Offset[numRow - 1, numCol - 1];
            Range writeRange = objSheet.Range[startCell, endCell];
            writeRange.Value2 = dataArray;

            objBook.Application.ScreenUpdating = true;
            objSheet = null;
            MessageBox.Show("Coding completed successfully", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CopyElements(ETABSv1.cSapModel SapModel, string currentGroup, string currentLabel, double targX, double targY, double refX, double refY, double rot, string mirr, bool groupAsUnit = true, bool groupAsElement = true, string latestGroupName = "..Last Duplicated", double dZ = 0)
        {
            int ret = 0;

            // Check if need to add to latest group
            bool groupAsLatest = true;
            if (latestGroupName == "")
            {
                groupAsLatest = false;
            }

            // Get objects in currentGroup
                int NumberItems = 0;
            int[] ObjectType = new int[1];
            string[] ObjectName = new string[1];
            ret = SapModel.GroupDef.GetAssignments(currentGroup, ref NumberItems, ref ObjectType, ref ObjectName);

            // For each object in the group 
            string unitNm = currentLabel + ".";
            string nameMod = "." + unitNm;
            string newName_J = "";

            for (int i = 0; i < NumberItems; i++)
            {
                switch (ObjectType[i])
                {
                    case 1: //Point
                        // Get coordinate data for joint
                        double x_J = 0;
                        double y_J = 0;
                        double z_J = 0;
                        ret = SapModel.PointObj.GetCoordCartesian(ObjectName[i], ref x_J, ref y_J, ref z_J);

                        // Calculate position of new coordinate
                        (double xFinal_J, double yFinal_J, double zFinal_J) = CalculateNewCoordinates(x_J, y_J, z_J, targX, targY, dZ, refX, refY, rot, mirr);

                        // Write new coordinate
                        newName_J = nameMod + ObjectName[i];
                        string finalName_J = "";
                        ret = SapModel.PointObj.AddCartesian(xFinal_J, yFinal_J, zFinal_J, ref finalName_J, newName_J);
                        // Need to copy over joint assignments
                        
                        // Assign joint restraint
                        bool[] restraint_J = new bool[6];
                        ret = SapModel.PointObj.GetRestraint(ObjectName[i], ref restraint_J);
                        ret = SapModel.PointObj.SetRestraint(finalName_J, ref restraint_J);

                        // Read joint load
                        int NumberPLoads_J = -1;
                        string[] PointName_J = new string[0];
                        string[] LoadPat_J = new string[0];
                        int[] LCStep_J = new int[0];
                        string[] CSys_J = new string[0];
                        double[] F1_J = new double[0];
                        double[] F2_J = new double[0];
                        double[] F3_J = new double[0];
                        double[] M1_J = new double[0];
                        double[] M2_J = new double[0];
                        double[] M3_J = new double[0];

                        ret = SapModel.PointObj.GetLoadForce(ObjectName[i], ref NumberPLoads_J, ref PointName_J, ref LoadPat_J, ref LCStep_J, ref CSys_J, ref F1_J, ref F2_J, ref F3_J, ref M1_J, ref M2_J, ref M3_J);
                        double[] LoadValue_J = new double[6];
                        // Rotate and assign joint loads
                        for (int j = 0; j < NumberPLoads_J; j++)
                        {
                            if ((rot == 0) && (mirr != "X") && (mirr != "Y"))
                            {
                                LoadValue_J[0] = F1_J[j];
                                LoadValue_J[1] = F2_J[j];
                                LoadValue_J[2] = F3_J[j];
                                LoadValue_J[3] = M1_J[j];
                                LoadValue_J[4] = M2_J[j];
                                LoadValue_J[5] = M3_J[j];
                            }
                            else
                            {
                                (LoadValue_J[0], LoadValue_J[1], LoadValue_J[2], LoadValue_J[3], LoadValue_J[4], LoadValue_J[5]) = RotateJointLoad(F1_J[j], F2_J[j], F3_J[j], M1_J[j], M2_J[j], M3_J[j], rot, mirr);
                            }
                            ret = SapModel.PointObj.SetLoadForce(finalName_J, LoadPat_J[j], ref LoadValue_J, false, CSys_J[j]);
                        }

                        // Assign to Unit Group 
                        if (groupAsUnit)
                        {
                            string GrpName = "." + currentGroup;
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.PointObj.SetGroupAssign(finalName_J, GrpName);
                        }
                        if (groupAsLatest)
                        {
                            ret = SapModel.PointObj.SetGroupAssign(finalName_J, latestGroupName);
                        }
                        break;

                    case 2: //Frame
                        // Get frame data
                        string[] nodes = new string[2];
                        ret = SapModel.FrameObj.GetPoints(ObjectName[i], ref nodes[0], ref nodes[1]);

                        // Get coordinates from point names and calculate final position
                        double[] x_F = new double[2];
                        double[] y_F = new double[2];
                        double[] z_F = new double[2];
                        double[] xFinal_F = new double[2];
                        double[] yFinal_F = new double[2];
                        double[] zFinal_F = new double[2];

                        for (int j = 0; j < 2; j++)
                        {
                            ret = SapModel.PointObj.GetCoordCartesian(nodes[j], ref x_F[j], ref y_F[j], ref z_F[j]);
                            (xFinal_F[j], yFinal_F[j], zFinal_F[j]) = CalculateNewCoordinates(x_F[j], y_F[j], z_F[j], targX, targY, dZ, refX, refY, rot, mirr);
                        }

                        // Check if start and stop coordinates have shifted relative x
                        bool anglecheck = CheckRelativeNodes(x_F, y_F, xFinal_F, yFinal_F);

                        // Get section type 
                        string PropName_F = "";
                        string SAuto_F = "";
                        ret = SapModel.FrameObj.GetSection(ObjectName[i], ref PropName_F, ref SAuto_F);

                        // Add frame
                        string newName_F = nameMod + ObjectName[i]+".";
                        if (!anglecheck)
                        {
                            newName_F = newName_F + "R";
                        }
                        string finalName_F = "";
                        ret = SapModel.FrameObj.AddByCoord(xFinal_F[0], yFinal_F[0], zFinal_F[0], xFinal_F[1], yFinal_F[1], zFinal_F[1], ref finalName_F, PropName_F, newName_F);

                        // Assign Local Axis
                        double Ang = 0;
                        bool Advanced = false;
                        ret = SapModel.FrameObj.GetLocalAxes(ObjectName[i], ref Ang, ref Advanced);
                        if (z_F[0] != z_F[1]) // Find Column
                        {
                            Ang = Ang + rot;
                            if (mirr == "Y")
                            {
                                Ang = Ang + 180;
                            }
                        }
                        ret = SapModel.FrameObj.SetLocalAxes(finalName_F, Ang);

                        // Assign Insert Point
                        int CardinalPoint = 0;
                        bool Mirror2 = false;
                        bool Mirror3 = false;
                        bool StiffTransform = false;
                        double[] Offset1 = null;
                        double[] Offset2 = null;
                        string CSys = null;
                        ret = SapModel.FrameObj.GetInsertionPoint_1(ObjectName[i], ref CardinalPoint, ref Mirror2, ref Mirror3, ref StiffTransform, ref Offset1, ref Offset2, ref CSys);
                        ret = SapModel.FrameObj.SetInsertionPoint_1(finalName_F, CardinalPoint, Mirror2, Mirror3, StiffTransform, ref Offset1, ref Offset2, CSys);

                        // Assign End Length Offsets 
                        bool AutoOffset = true;
                        double Length1 = 0.0;
                        double Length2 = 0.0;
                        double RZ = 0.0;
                        ret = SapModel.FrameObj.GetEndLengthOffset(ObjectName[i], ref AutoOffset, ref Length1, ref Length2, ref RZ);
                        ret = SapModel.FrameObj.SetEndLengthOffset(finalName_F, AutoOffset, Length1, Length2, RZ);
                        
                        // Assign Distributed Load
                        int LoadCount = 0;
                        string[] FrameName = new string[0];
                        string[] LoadPatF = new string[0];
                        int[] MyType = new int[0];
                        string[] CSysF = new string[0];
                        int[] Dir = new int[0];
                        double[] RD1 = new double[0];
                        double[] RD2 = new double[0];
                        double[] Dist1 = new double[0];
                        double[] Dist2 = new double[0];
                        double[] Val1 = new double[0];
                        double[] Val2 = new double[0];

                        ret = SapModel.FrameObj.GetLoadDistributed(ObjectName[i], ref LoadCount, ref FrameName, ref LoadPatF, ref MyType, ref CSysF, ref Dir, ref RD1, ref RD2, ref Dist1, ref Dist2, ref Val1, ref Val2);
 
                        if (!anglecheck) // To flip load assign if local axis is rotated
                        {
                            double[] RD1Final = RD2;
                            double[] RD2Final = RD1;
                            for (int j = 0; j<LoadCount; j++)
                            {
                                RD1[j] = 1 - RD1[j];
                                RD2[j] = 1 - RD2[j];
                            }
                        }
                        else
                        {
                            double[] RD1Final = RD1;
                            double[] RD2Final = RD2;
                        }

                        for (int j = 0; j < LoadCount; j++)
                        {
                            ret = SapModel.FrameObj.SetLoadDistributed(finalName_F, LoadPatF[j], MyType[j], Dir[j], RD1[j], RD2[j], Val1[j], Val2[j], CSysF[j], true, false); // 1st true is RelDist, 2nd false is whether to replace
                        }

                        // Assign Point Load
                        double[] RelDist = new double[0];
                        double[] Dist = new double[0];
                        double[] Val = new double[0];
                        ret = SapModel.FrameObj.GetLoadPoint(ObjectName[i], ref LoadCount, ref FrameName, ref LoadPatF, ref MyType, ref CSysF, ref Dir, ref RelDist, ref Dist, ref Val);

                        for (int j = 0; j < LoadCount; j++)
                        {
                            if (!anglecheck) // To flip load assign if local axis is rotated
                            {
                                RelDist[j] = 1 - RelDist[j];
                            }
                            ret = SapModel.FrameObj.SetLoadPoint(finalName_F, LoadPatF[j], MyType[j], Dir[j], RelDist[j], Val[j], CSysF[j], true, false); // 1st true is RelDist, 2nd false is whether to replace
                        }

                        // Assign Releases
                        bool[] II = new bool[0];
                        bool[] JJ = new bool[0];
                        double[] StartValue = new double[0];
                        double[] EndValue = new double[0];
                        ret = SapModel.FrameObj.GetReleases(ObjectName[i], ref II, ref JJ, ref StartValue, ref EndValue);
                        if (!anglecheck) // To flip load assign if local axis is rotated
                        {
                            ret = SapModel.FrameObj.SetReleases(finalName_F, ref JJ, ref II, ref EndValue, ref StartValue); // Swap start and end
                        }
                        else
                        {
                            ret = SapModel.FrameObj.SetReleases(finalName_F, ref II, ref JJ, ref StartValue, ref EndValue);
                        }
                        

                        // Assign Modifiers
                        double[] Value = new double[0];
                        ret = SapModel.FrameObj.GetModifiers(ObjectName[i], ref Value);
                        ret = SapModel.FrameObj.SetModifiers(finalName_F, ref Value);

                        // Assign to group
                        if (groupAsElement)
                        {
                            string GrpName = ".F." + ObjectName[i];
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.FrameObj.SetGroupAssign(finalName_F, GrpName);
                        }
                        if (groupAsUnit)
                        {
                            string GrpName = "." + currentGroup;
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.FrameObj.SetGroupAssign(finalName_F, GrpName);
                        }
                        if (groupAsLatest)
                        {
                            ret = SapModel.FrameObj.SetGroupAssign(finalName_F, latestGroupName);
                        }
                        break;

                    case 3: //Cable
                        break;

                    case 4: //Tendon
                        break;

                    case 5: //Area
                        // Get area data
                        int NumberPoints = -1;
                        string[] Point = new string[0];
                        ret = SapModel.AreaObj.GetPoints(ObjectName[i], ref NumberPoints, ref Point);

                        // Get coordinates from point names
                        double[] xList_S = new double[NumberPoints];
                        double[] yList_S = new double[NumberPoints];
                        double[] zList_S = new double[NumberPoints];

                        for (int j = 0; j < NumberPoints; j++)
                        {
                            double xIndv_S = 0;
                            double yIndv_S = 0;
                            double zIndv_S = 0;
                            ret = SapModel.PointObj.GetCoordCartesian(Point[j], ref xIndv_S, ref yIndv_S, ref zIndv_S);
                            (xList_S[j], yList_S[j], zList_S[j]) = CalculateNewCoordinates(xIndv_S, yIndv_S, zIndv_S, targX, targY, dZ, refX, refY, rot, mirr);
                        }

                        // Get Properties
                        string PropName_S = "";
                        ret = SapModel.AreaObj.GetProperty(ObjectName[i], ref PropName_S);

                        // Add area
                        newName_J = nameMod + ObjectName[i];
                        string finalName_S = "";
                        ret = SapModel.AreaObj.AddByCoord(NumberPoints, ref xList_S, ref yList_S, ref zList_S, ref finalName_S, PropName_S, newName_J);

                        // Assign Pier Label
                        string PierName = "";
                        ret = SapModel.AreaObj.GetPier(ObjectName[i], ref PierName);
                        if (PierName != "None") // Only add pier label if it has already been assigned
                        {
                            PierName = nameMod + PierName;
                            ret = SapModel.PierLabel.SetPier(PierName);
                            ret = SapModel.AreaObj.SetPier(finalName_S, PierName);
                        }

                        // Get Uniform Load
                        int NumberItems_SL = -1;
                        string[] AreaName_SL = new string[0];
                        string[] LoadPat_SL = new string[0];
                        string[] CSys_SL = new string[0];
                        int[] Dir_SL = new int[0];
                        double[] Value_SL = new double[0];

                        ret = SapModel.AreaObj.GetLoadUniform(ObjectName[i], ref NumberItems_SL, ref AreaName_SL, ref LoadPat_SL, ref CSys_SL, ref Dir_SL, ref Value_SL);
                        for (int j = 0; j < NumberItems_SL; j++)
                        {
                            ret = SapModel.AreaObj.SetLoadUniform(finalName_S, LoadPat_SL[j], Value_SL[j], Dir_SL[j], false, CSys_SL[j]);
                        }

                        // Assign to group
                        if (groupAsElement)
                        {
                            string GrpName = ".S." + ObjectName[i];
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.AreaObj.SetGroupAssign(finalName_S, GrpName);
                        }
                        if (groupAsUnit)
                        {
                            string GrpName = "." + currentGroup;
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.AreaObj.SetGroupAssign(finalName_S, GrpName);
                        }
                        if (groupAsLatest)
                        {
                            ret = SapModel.AreaObj.SetGroupAssign(finalName_S, latestGroupName);
                        }
                        break;

                    case 6: //Solid
                        break;

                    case 7: //Link
                        // Get Link Data
                        string Point1 = "";
                        string Point2 = "";
                        ret = SapModel.LinkObj.GetPoints(ObjectName[i], ref Point1, ref Point2);

                        // Get coordinate of points
                        double[] Point1Coord = new double[3]; // x, y, z
                        double[] Point2Coord = new double[3];
                        double[] Point1Coord_final = new double[3]; // coordinate after transformation
                        double[] Point2Coord_final = new double[3];
                        bool boolIsSingleJoint = false;

                        if (Point1 == Point2)
                        {
                            boolIsSingleJoint = true;
                            ret = SapModel.PointObj.GetCoordCartesian(Point1, ref Point1Coord[0], ref Point1Coord[1], ref Point1Coord[2]);
                            (Point1Coord_final[0], Point1Coord_final[1], Point1Coord_final[2]) = CalculateNewCoordinates(Point1Coord[0], Point1Coord[1], Point1Coord[2], targX, targY, dZ, refX, refY, rot, mirr);
                        }
                        else
                        { 
                            // Get coordinate data for joint
                            ret = SapModel.PointObj.GetCoordCartesian(Point1, ref Point1Coord[0], ref Point1Coord[1], ref Point1Coord[2]);
                            ret = SapModel.PointObj.GetCoordCartesian(Point2, ref Point2Coord[0], ref Point2Coord[1], ref Point2Coord[2]);

                            // Calculate position of new coordinate
                            (Point1Coord_final[0], Point1Coord_final[1], Point1Coord_final[2]) = CalculateNewCoordinates(Point1Coord[0], Point1Coord[1], Point1Coord[2], targX, targY, dZ, refX, refY, rot, mirr);
                            (Point2Coord_final[0], Point2Coord_final[1], Point2Coord_final[2]) = CalculateNewCoordinates(Point2Coord[0], Point2Coord[1], Point2Coord[2], targX, targY, dZ, refX, refY, rot, mirr);
                        }

                        // Get link property
                        string PropNameLink = "";
                        ret = SapModel.LinkObj.GetProperty(ObjectName[i],ref PropNameLink);
                        // Add new link
                        string newNameLink = nameMod + ObjectName[i];
                        string finalName_L = "";
                        ret = SapModel.LinkObj.AddByCoord(Point1Coord_final[0], Point1Coord_final[1], Point1Coord_final[2], Point2Coord_final[0], Point2Coord_final[1], Point2Coord_final[2], ref finalName_L, boolIsSingleJoint, PropNameLink, newNameLink);

                        // Assign to group
                        if (groupAsElement)
                        {
                            string GrpName = ".L." + ObjectName[i]; 
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.LinkObj.SetGroupAssign(finalName_L, GrpName);
                        }
                        if (groupAsUnit)
                        {
                            string GrpName = "." + currentGroup;
                            ret = SapModel.GroupDef.SetGroup(GrpName);
                            ret = SapModel.LinkObj.SetGroupAssign(finalName_L, GrpName);
                        }
                        if (groupAsLatest)
                        {
                            ret = SapModel.LinkObj.SetGroupAssign(finalName_L, latestGroupName);
                        }
                        break;

                    default:
                        Console.WriteLine("Warning: Unidentified Object");
                        break;
                }
            }
            
        }

        private (double, double, double) CalculateNewCoordinates(double x, double y, double z, double targX, double targY, double dZ, double refX, double refY, double rot, string mirr)
        {
            // Calculate position of new coordinate
            double xFinal = x;
            double yFinal = y;
            rot = rot * (Math.PI / 180);

            // 1. Mirror 
            if (mirr == "Y")
            {
                xFinal = 2 * refX - x;
            }
            else if (mirr == "X")
            {
                yFinal = 2 * refY - y;
            }

            // 2. Rotation
            if ((rot != 0))
            {
                double xMirr = xFinal;
                double yMirr = yFinal;
                xFinal = refX + (xMirr - refX) * Math.Cos(rot) - (yMirr - refY) * Math.Sin(rot);
                yFinal = refY + (xMirr - refX) * Math.Sin(rot) + (yMirr - refY) * Math.Cos(rot);
            }

            // 3. Translate
            xFinal = Math.Round(xFinal + (targX - refX),4);
            yFinal = Math.Round(yFinal + (targY - refY),4);
            double zFinal = z + dZ;

            return (xFinal, yFinal, zFinal);
        }

        private bool CheckRelativeNodes(double[] xInitial, double[] yInitial, double[] xFinal, double[] yFinal)
        {
            double angleInitial = Math.Atan2(yInitial[1] - yInitial[0], xInitial[1] - xInitial[0]); // find original angle in rad
            double angleFinal = Math.Atan2(yFinal[1] - yFinal[0], xFinal[1] - xFinal[0]); // find new angle in rad

            bool angleTypeInitial = false; // anlgeType = true means node 1 pointing to node 2
            if (angleInitial <= Math.PI/2 && angleInitial > -Math.PI/2)
            {
                angleTypeInitial = true;
            }
            bool angleTypeFinal = false;

            if (angleFinal <= Math.PI / 2 && angleFinal > -Math.PI / 2)
            {
                angleTypeFinal = true;
            }
            bool angleTypeMatch = angleTypeFinal == angleTypeInitial;

            return angleTypeMatch;
        }

        private (double, double, double, double, double, double) RotateJointLoad(double Fx, double Fy, double Fz, double Mx, double My, double Mz, double rot, string mirr)
        {
            double Fx_mirr = Fx;
            double Fy_mirr = Fy;
            double Fz_mirr = Fz;
            double Mx_mirr = Mx;
            double My_mirr = My;
            double Mz_mirr = Mz;
            rot = rot * (Math.PI / 180); // convert to radians

            if (mirr == "X")
            {
                Fy_mirr = -Fy;
                My_mirr = -My;
                Mz_mirr = -Mz;
            }
            else if (mirr == "Y")
            {
                Fx_mirr = -Fx;
                Mx_mirr = -Mx;
                Mz_mirr = -Mz;
            }

            if (rot == 0)
            {
                return (Fx_mirr, Fy_mirr, Fz_mirr, Mx_mirr, My_mirr, Mz_mirr);
            }
            else
            {
                double Fx_final = Fx_mirr;
                double Fy_final = Fy_mirr;
                double Fz_final = Fz_mirr;
                double Mx_final = Mx_mirr;
                double My_final = My_mirr;
                double Mz_final = Mz_mirr;

                Fx_final = Fx_mirr * Math.Cos(rot) - Fy_mirr * Math.Sin(rot);
                Fy_final = Fx_mirr * Math.Sin(rot) + Fy_mirr * Math.Cos(rot);
                Mx_final = Mx_mirr * Math.Cos(rot) - My_mirr * Math.Sin(rot);
                My_final = Mx_mirr * Math.Sin(rot) + My_mirr * Math.Cos(rot);

                return (Fx_final, Fy_final, Fz_final, Mx_final, My_final, Mz_final);
            }
            
            
        }


        private void EnableAllLC(ETABSv1.cSapModel mySapModel)
        {
            int ret2 = 0;
            //// Get all load case
            int NumberNames = 0;
            string[] MyName = null;
            ret2 = mySapModel.LoadCases.GetNameList(ref NumberNames, ref MyName);
            // Enable all load case
            for (int i = 0; i < NumberNames; i++)
            {
                ret2 = mySapModel.Results.Setup.SetCaseSelectedForOutput(MyName[i]);
                bool sel = false;
                ret2 = mySapModel.Results.Setup.GetCaseSelectedForOutput(MyName[i], ref sel);
            }


            //// Get all load combi
            int NumberNames2 = 0;
            string[] MyName2 = null;
            ret2 = mySapModel.RespCombo.GetNameList(ref NumberNames2, ref MyName2);
            for (int i = 0; i < NumberNames2; i++)
            {
                ret2 = mySapModel.Results.Setup.SetComboSelectedForOutput(MyName2[i]);
                bool sel2 = false;
                ret2 = mySapModel.Results.Setup.GetCaseSelectedForOutput(MyName2[i], ref sel2);
            }
        }
        #endregion


        string lastSelectedRange = null;

        private void Test_Click(object sender, RibbonControlEventArgs e)
        {
            
            Range selectedRange;
            if (string.IsNullOrEmpty(lastSelectedRange))
            {
                selectedRange = Globals.ThisAddIn.Application.InputBox("Select range for SetRangeofGetJointReactionZ", "Select reference title range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                lastSelectedRange = selectedRange.get_Address(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                //lastSelectedRange = rangeString;
            }
            else
            {
                selectedRange = Globals.ThisAddIn.Application.Range[lastSelectedRange];
            }

            Workbook ThisWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Microsoft.Office.Core.DocumentProperties properties = ThisWorkBook.CustomDocumentProperties;
            
            // Check if custom property exist, replace value if exist
            Boolean propertyExist = false;
            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == "Active Range")
                {
                    prop.Value = selectedRange.Address;
                    propertyExist = true;
                    break;
                }
            }
            // If doesn't exist, create and set value
            if (!propertyExist)
            {
                properties.Add("Active Range", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, selectedRange.Address[false, false]);
            }

            // Test Read
            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == "Active Range")
                {
                    MessageBox.Show("Successfully added " + prop.Value, "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            lastSelectedRange = selectedRange.Address[false, false];
        }

        private void GetPierTest_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook ThisWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Office.DocumentProperties properties = ThisWorkBook.CustomDocumentProperties;

            Boolean propertyExist = false;
            foreach (Office.DocumentProperty prop in ThisWorkBook.CustomDocumentProperties)
            {
                if (prop.Name == "Active Range")
                {
                    MessageBox.Show(prop.Value, "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    propertyExist = true;
                    break;
                }
            }
            if (!propertyExist)
            {
                MessageBox.Show("Active Range not set", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //MessageBox.Show("Test " + lastSelectedRange, "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GetJointTest_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook ThisWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Office.DocumentProperties properties = ThisWorkBook.CustomDocumentProperties;

            Boolean propertyExist = false;
            foreach (Office.DocumentProperty prop in ThisWorkBook.CustomDocumentProperties)
            {
                if (prop.Name == "Active Range")
                {
                    MessageBox.Show(prop.Value, "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    propertyExist = true;
                    break;
                }
            }

            if (!propertyExist)
            {
                MessageBox.Show("Active Range not set", "PWG_Meinhardt Automation Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AutocadTest_Click(object sender, RibbonControlEventArgs e)
        {

        }


    }
}
