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
using System.Runtime.InteropServices;
using System.IO;

namespace ExcelAddIn2
{
    public partial class ETABSTaskPane : UserControl
    {
        public ETABSTaskPane()
        {
            InitializeComponent();
        }
        #region Format Tools
        #region Unit Duplicator

        private void getGroups_Click(object sender, EventArgs e)
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

            MessageBox.Show("Completed", "Completed");
        }

        private void getSelCoord_Click(object sender, EventArgs e)
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
            CommonUtilities.WriteToExcel(0, 0, false ,selectedJoints2, x2, y2, z2);
            MessageBox.Show("Completed", "Completed");
        }

        private void getFloors_Click(object sender, EventArgs e)
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
            MessageBox.Show("Completed", "Completed");
        }
        #endregion

        #region Shared 
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


        #endregion

        #region Utlities
        private void checkWalls_Click(object sender, EventArgs e)
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
            MessageBox.Show("Completed", "Completed");
        }

        private void selectBeamLabel_Click(object sender, EventArgs e)
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
            MessageBox.Show("Completed", "Completed");
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

        private void selectErrorJoint_Click(object sender, EventArgs e)
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
            string errorFileName = filePath.Substring(0, filePath.Length - 4) + ".LOG";
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

            CommonUtilities.WriteToExcel(0, 0, false, errorJoints.ToArray(), coord1.ToArray(), coord2.ToArray(), coord3.ToArray(), grouped.ToArray());
            string msgText = "Coding completed, " + counter.ToString() + " added.";
            MessageBox.Show(msgText, "Completed");
        }

        #endregion
        #endregion

        #region Comaprison Tab
        
        #endregion
    }
}
