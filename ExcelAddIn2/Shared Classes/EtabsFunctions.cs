using ETABSv1;
using ExcelAddIn2.Excel_Pane_Folder.HDB_Design;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace ExcelAddIn2
{
    public static class EtabsFunctions
    {
        #region Replicate Element
        public static void ReplicateSelectedElement(cOAPI etabsObject, cSapModel sapModel, double[,] offsetValues)
        {
            (int numSel, int[] objType, string[] objName) = GetSelectedElements(sapModel);
            for (int objNum = 0; objNum < numSel; objNum++)
            {
                for (int rowNum = 0; rowNum < offsetValues.GetLength(0); rowNum++)
                {
                    double[] offsetValue = new double[3] { offsetValues[rowNum, 0], offsetValues[rowNum, 1], offsetValues[rowNum, 2] };
                    CopyElement(sapModel, objType[objNum], objName[objNum], offsetValue);
                }
            }
        }
        #endregion

        #region ETABS Common
        public static (int numSel, int[] objectType, string[] ObjectName) GetSelectedElements(cSapModel sapModel)
        {
            int ret = 0;
            int NumSel = 0;
            int[] ObjectType = new int[0];
            string[] ObjectName = new string[0];
            ret = sapModel.SelectObj.GetSelected(ref NumSel, ref ObjectType, ref ObjectName);
            return (NumSel, ObjectType, ObjectName);
        }

        public static (int numSel, string[] ObjectName) GetSelectedElementsByType(cSapModel sapModel, EtabsObjectType type)
        {
            // come back to refractor. Need to add function to get all only
            (int numSelAll, int[] objTypeAll, string[] objNameAll) = GetSelectedElements(sapModel);

            List<string> wallUNs = new List<string>();
            int wallCount = 0;

            for (int i = 0; i < numSelAll; i++)
            {
                if (objTypeAll[i] != (int)type) { continue; }

                wallUNs.Add(objNameAll[i]);
                wallCount++;
            }

            return (wallCount, wallUNs.ToArray());
        }

        static void CopyElement(cSapModel SapModel, int objType, string objName, double[] offsetValue)
        {
            // Missing parameters that we assume to be 0
            double rot = 0;
            string mirr = "";

            // This is from old code, to be refractored
            int ret = 0;

            switch (objType)
            {
                case 1: //Point
                    {
                        #region Adding new joint
                        // Get coordinate data for joint
                        double x = 0;
                        double y = 0;
                        double z = 0;
                        ret = SapModel.PointObj.GetCoordCartesian(objName, ref x, ref y, ref z);

                        // Calculate position of new coordinate
                        double xFinal = x + offsetValue[0];
                        double yFinal = y + offsetValue[1];
                        double zFinal = z + offsetValue[2];

                        // Add new coordinate
                        string newJointName = "";
                        ret = SapModel.PointObj.AddCartesian(xFinal, yFinal, zFinal, ref newJointName);
                        #endregion

                        #region Copying settings to New Joint
                        // Assign joint restraint
                        bool[] restraint = new bool[6];
                        ret = SapModel.PointObj.GetRestraint(objName, ref restraint);
                        ret = SapModel.PointObj.SetRestraint(newJointName, ref restraint);

                        // Read joint load
                        int NumberPLoads = -1;
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

                        ret = SapModel.PointObj.GetLoadForce(objName, ref NumberPLoads, ref PointName, ref LoadPat, ref LCStep, ref CSys, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
                        double[] LoadValue_J = new double[6];

                        // Rotate and assign joint loads
                        for (int j = 0; j < NumberPLoads; j++)
                        {
                            if ((rot == 0) && (mirr != "X") && (mirr != "Y"))
                            {
                                LoadValue_J[0] = F1[j];
                                LoadValue_J[1] = F2[j];
                                LoadValue_J[2] = F3[j];
                                LoadValue_J[3] = M1[j];
                                LoadValue_J[4] = M2[j];
                                LoadValue_J[5] = M3[j];
                            }
                            else
                            {
                                (LoadValue_J[0], LoadValue_J[1], LoadValue_J[2], LoadValue_J[3], LoadValue_J[4], LoadValue_J[5]) = RotateJointLoad(F1[j], F2[j], F3[j], M1[j], M2[j], M3[j], rot, mirr);
                            }
                            ret = SapModel.PointObj.SetLoadForce(newJointName, LoadPat[j], ref LoadValue_J, false, CSys[j]);
                        }
                        #endregion
                        break;
                    }
                case 2: //Frame
                    {
                        #region Get new coordinates for new frame
                        // Get frame data
                        string[] jointNames = new string[2];
                        ret = SapModel.FrameObj.GetPoints(objName, ref jointNames[0], ref jointNames[1]);

                        // Calculate new position of joints
                        int numJoints = 2;
                        double[,] ogCoord = new double[numJoints, 3];
                        double[,] finalCoord = new double[numJoints, 3];

                        // Get coordinates from point names and calculate final position
                        for (int jointNum = 0; jointNum < numJoints; jointNum++)
                        {
                            ret = SapModel.PointObj.GetCoordCartesian(jointNames[jointNum], ref ogCoord[jointNum, 0], ref ogCoord[jointNum, 1], ref ogCoord[jointNum, 2]);
                            for (int coordNum = 0; coordNum < 3; coordNum++)
                            {
                                finalCoord[jointNum, coordNum] = ogCoord[jointNum, coordNum] + offsetValue[coordNum];
                            }
                        }
                        #endregion

                        #region Get Some Properties Before Creating
                        // Check rotation, To refractor this
                        // Check if start and stop coordinates have shifted relative x
                        bool frameFlipped = CheckRelativeNodes(
                            new double[] { ogCoord[0, 0], ogCoord[1, 0] },
                            new double[] { ogCoord[0, 1], ogCoord[1, 1] },
                            new double[] { finalCoord[0, 0], ogCoord[1, 0] },
                            new double[] { finalCoord[0, 1], ogCoord[1, 1] }
                            );                         // x, y, xFinal, yFinal

                        // Get section type 
                        string PropName = "";
                        string SAuto = "";
                        ret = SapModel.FrameObj.GetSection(objName, ref PropName, ref SAuto);
                        #endregion

                        #region Create new frame
                        //string newName_F = nameMod + ObjectName[i] + ".";
                        //if (!frameFlipped)
                        //{
                        //    newName_F = newName_F + "R";
                        //}
                        string finalName_F = "";
                        ret = SapModel.FrameObj.AddByCoord(
                            finalCoord[0, 0], finalCoord[0, 1], finalCoord[0, 2],
                            finalCoord[1, 0], finalCoord[1, 1], finalCoord[1, 2],
                            ref finalName_F, PropName);
                        #endregion

                        #region Copy Settings to new frame
                        // Assign Local Axis
                        double Ang = 0;
                        bool Advanced = false;
                        ret = SapModel.FrameObj.GetLocalAxes(objName, ref Ang, ref Advanced);
                        if (ogCoord[0, 2] != ogCoord[1, 2]) // Find Column
                        {
                            // Handle rotation for column
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
                        ret = SapModel.FrameObj.GetInsertionPoint_1(objName, ref CardinalPoint, ref Mirror2, ref Mirror3, ref StiffTransform, ref Offset1, ref Offset2, ref CSys);
                        ret = SapModel.FrameObj.SetInsertionPoint_1(finalName_F, CardinalPoint, Mirror2, Mirror3, StiffTransform, ref Offset1, ref Offset2, CSys);

                        // Assign End Length Offsets 
                        bool AutoOffset = true;
                        double Length1 = 0.0;
                        double Length2 = 0.0;
                        double RZ = 0.0;
                        ret = SapModel.FrameObj.GetEndLengthOffset(objName, ref AutoOffset, ref Length1, ref Length2, ref RZ);
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

                        ret = SapModel.FrameObj.GetLoadDistributed(objName, ref LoadCount, ref FrameName, ref LoadPatF, ref MyType, ref CSysF, ref Dir, ref RD1, ref RD2, ref Dist1, ref Dist2, ref Val1, ref Val2);

                        if (!frameFlipped) // To flip load assign if local axis is rotated
                        {
                            for (int j = 0; j < LoadCount; j++)
                            {
                                RD1[j] = 1 - RD1[j];
                                RD2[j] = 1 - RD2[j];
                            }
                        }

                        for (int j = 0; j < LoadCount; j++)
                        {
                            ret = SapModel.FrameObj.SetLoadDistributed(finalName_F, LoadPatF[j], MyType[j], Dir[j], RD1[j], RD2[j], Val1[j], Val2[j], CSysF[j], true, false); // 1st true is RelDist, 2nd false is whether to replace
                        }

                        // Assign Point Load
                        double[] RelDist = new double[0];
                        double[] Dist = new double[0];
                        double[] Val = new double[0];
                        ret = SapModel.FrameObj.GetLoadPoint(objName, ref LoadCount, ref FrameName, ref LoadPatF, ref MyType, ref CSysF, ref Dir, ref RelDist, ref Dist, ref Val);

                        for (int j = 0; j < LoadCount; j++)
                        {
                            if (!frameFlipped) // To flip load assign if local axis is rotated
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
                        ret = SapModel.FrameObj.GetReleases(objName, ref II, ref JJ, ref StartValue, ref EndValue);

                        if (!frameFlipped) // To flip load assign if local axis is rotated
                        {
                            ret = SapModel.FrameObj.SetReleases(finalName_F, ref JJ, ref II, ref EndValue, ref StartValue); // Swap start and end
                        }
                        else
                        {
                            ret = SapModel.FrameObj.SetReleases(finalName_F, ref II, ref JJ, ref StartValue, ref EndValue);
                        }

                        // Assign Modifiers
                        double[] Value = new double[0];
                        ret = SapModel.FrameObj.GetModifiers(objName, ref Value);
                        ret = SapModel.FrameObj.SetModifiers(finalName_F, ref Value);
                        #endregion
                        break;
                    }

                case 3: //Cable
                    throw new Exception("Cable replication not yet implemented");

                case 4: //Tendon
                    throw new Exception("Tendon replication not yet implemented");

                case 5: //Area
                    {
                        #region Get new coordinates for new area
                        // Get area data
                        int numJoints = -1;
                        string[] jointNames = new string[0];
                        ret = SapModel.AreaObj.GetPoints(objName, ref numJoints, ref jointNames);

                        // Calculate new position of joints
                        double[,] ogCoord = new double[numJoints, 3];
                        // Area handles coordinates differently from frame
                        //double[,] finalCoord = new double[numJoints, 3]; 
                        double[] xFinal = new double[numJoints];
                        double[] yFinal = new double[numJoints];
                        double[] zFinal = new double[numJoints];



                        // Get coordinates from point names and calculate final position
                        for (int jointNum = 0; jointNum < numJoints; jointNum++)
                        {
                            ret = SapModel.PointObj.GetCoordCartesian(jointNames[jointNum], ref ogCoord[jointNum, 0], ref ogCoord[jointNum, 1], ref ogCoord[jointNum, 2]);
                            xFinal[jointNum] = ogCoord[jointNum, 0] + offsetValue[0];
                            yFinal[jointNum] = ogCoord[jointNum, 1] + offsetValue[1];
                            zFinal[jointNum] = ogCoord[jointNum, 2] + offsetValue[2];
                        }

                        //// Get coordinates from point names
                        //double[] xList_S = new double[numJoints];
                        //double[] yList_S = new double[numJoints];
                        //double[] zList_S = new double[numJoints];

                        //for (int j = 0; j < numJoints; j++)
                        //{
                        //    double xIndv_S = 0;
                        //    double yIndv_S = 0;
                        //    double zIndv_S = 0;
                        //    ret = SapModel.PointObj.GetCoordCartesian(jointNames[j], ref xIndv_S, ref yIndv_S, ref zIndv_S);
                        //    (xList_S[j], yList_S[j], zList_S[j]) = CalculateNewCoordinates(xIndv_S, yIndv_S, zIndv_S, targX, targY, dZ, refX, refY, rot, mirr);
                        //}
                        #endregion

                        #region Get properties before creating
                        // Get Properties
                        string PropName = "";
                        ret = SapModel.AreaObj.GetProperty(objName, ref PropName);
                        #endregion

                        #region Add Area
                        // Add area
                        string finalName_S = "";
                        ret = SapModel.AreaObj.AddByCoord(numJoints, ref xFinal, ref yFinal, ref zFinal, ref finalName_S, PropName);
                        #endregion

                        #region Copy Settings
                        // Assign Pier Label
                        string PierName = "";
                        ret = SapModel.AreaObj.GetPier(objName, ref PierName);
                        ret = SapModel.AreaObj.SetPier(finalName_S, PierName);

                        // Get Uniform Load
                        int NumberItems = -1;
                        string[] AreaName = new string[0];
                        string[] LoadPat = new string[0];
                        string[] CSys = new string[0];
                        int[] Dir = new int[0];
                        double[] Value = new double[0];

                        ret = SapModel.AreaObj.GetLoadUniform(objName, ref NumberItems, ref AreaName, ref LoadPat, ref CSys, ref Dir, ref Value);
                        for (int j = 0; j < NumberItems; j++)
                        {
                            ret = SapModel.AreaObj.SetLoadUniform(finalName_S, LoadPat[j], Value[j], Dir[j], false, CSys[j]);
                        }
                        #endregion
                        break;
                    }

                case 6: //Solid
                    throw new Exception("Solid replication not yet implemented");

                case 7: //Link
                    {
                        #region Get new coordinates for new link
                        // Get Link Data
                        string Point1 = "";
                        string Point2 = "";
                        ret = SapModel.LinkObj.GetPoints(objName, ref Point1, ref Point2);


                        // Check if is a single joint link
                        bool boolIsSingleJoint = false;
                        int numJoints = 2;
                        string[] jointNames = new string[2] { Point1, Point2 };
                        if (Point1 == Point2) { boolIsSingleJoint = true; }

                        // Calculate new position of joints
                        double[,] ogCoord = new double[numJoints, 3];
                        double[,] finalCoord = new double[numJoints, 3];

                        // Get coordinates from point names and calculate final position
                        for (int jointNum = 0; jointNum < numJoints; jointNum++)
                        {
                            ret = SapModel.PointObj.GetCoordCartesian(jointNames[jointNum], ref ogCoord[jointNum, 0], ref ogCoord[jointNum, 1], ref ogCoord[jointNum, 2]);
                            for (int coordNum = 0; coordNum < 3; coordNum++)
                            {
                                finalCoord[jointNum, coordNum] = ogCoord[jointNum, coordNum] + offsetValue[coordNum];
                            }
                        }
                        #endregion

                        //// Get coordinate of points
                        //double[] Point1Coord = new double[3]; // x, y, z
                        //double[] Point2Coord = new double[3];
                        //double[] Point1Coord_final = new double[3]; // coordinate after transformation
                        //double[] Point2Coord_final = new double[3];
                        //if (Point1 == Point2)
                        //{
                        //    boolIsSingleJoint = true;
                        //    ret = SapModel.PointObj.GetCoordCartesian(Point1, ref Point1Coord[0], ref Point1Coord[1], ref Point1Coord[2]);
                        //    (Point1Coord_final[0], Point1Coord_final[1], Point1Coord_final[2]) = CalculateNewCoordinates(Point1Coord[0], Point1Coord[1], Point1Coord[2], targX, targY, dZ, refX, refY, rot, mirr);
                        //}
                        //else
                        //{
                        //    // Get coordinate data for joint
                        //    ret = SapModel.PointObj.GetCoordCartesian(Point1, ref Point1Coord[0], ref Point1Coord[1], ref Point1Coord[2]);
                        //    ret = SapModel.PointObj.GetCoordCartesian(Point2, ref Point2Coord[0], ref Point2Coord[1], ref Point2Coord[2]);

                        //    // Calculate position of new coordinate
                        //    (Point1Coord_final[0], Point1Coord_final[1], Point1Coord_final[2]) = CalculateNewCoordinates(Point1Coord[0], Point1Coord[1], Point1Coord[2], targX, targY, dZ, refX, refY, rot, mirr);
                        //    (Point2Coord_final[0], Point2Coord_final[1], Point2Coord_final[2]) = CalculateNewCoordinates(Point2Coord[0], Point2Coord[1], Point2Coord[2], targX, targY, dZ, refX, refY, rot, mirr);
                        //}

                        #region Add New Link
                        // Get link property
                        string PropNameLink = "";
                        ret = SapModel.LinkObj.GetProperty(objName, ref PropNameLink);
                        // Add new link
                        //string newNameLink = nameMod + ObjectName[i];
                        string finalName_L = "";
                        ret = SapModel.LinkObj.AddByCoord(
                            finalCoord[0, 0], finalCoord[0, 1], finalCoord[0, 2],
                            finalCoord[1, 0], finalCoord[1, 1], finalCoord[1, 2],
                            ref finalName_L, boolIsSingleJoint, PropNameLink);
                        #endregion
                        break;
                    }
                default:
                    throw new Exception("Object type not recognised");

            }

        }

        static bool CheckRelativeNodes(double[] xInitial, double[] yInitial, double[] xFinal, double[] yFinal)
        {
            double angleInitial = Math.Atan2(yInitial[1] - yInitial[0], xInitial[1] - xInitial[0]); // find original angle in rad
            double angleFinal = Math.Atan2(yFinal[1] - yFinal[0], xFinal[1] - xFinal[0]); // find new angle in rad

            bool angleTypeInitial = false; // anlgeType = true means node 1 pointing to node 2
            if (angleInitial <= Math.PI / 2 && angleInitial > -Math.PI / 2)
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

        static (double, double, double, double, double, double) RotateJointLoad(double Fx, double Fy, double Fz, double Mx, double My, double Mz, double rot, string mirr)
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
        #endregion

        #region ETABS Get Group Data
        static public (int[] objectTypeIds, string[] objectNames) GetGroupElements(cSapModel sapModel, string groupName)
        {
            int numberItems = 0;
            int[] objectTypeIDs = null;
            string[] objectNames = null;
            try
            {
                int ret = sapModel.GroupDef.GetAssignments(groupName, ref numberItems, ref objectTypeIDs, ref objectNames);
                if (ret != 0) { throw new Exception($"Failed to get elements assigned to group '{groupName}'. Ensure the group exists."); }
                return (objectTypeIDs, objectNames);
            }
            catch (Exception ex)
            {
                throw new Exception($"An unexpected error occurred while getting group elements: {ex.Message}");
            }
        }
        static public (int[] objectTypeIds, string[] objectNames) GetGroupElementofType(cSapModel sapModel, string groupName, HashSet<EtabsObjectType> targetObjectTypes)
        {
            List<int> objectTypeIdsListAfterFilter = new List<int>();
            List<string> objectNamesAfterFilter = new List<string>();
            
            (int[] objectTypeIds, string[] objectNames) = GetGroupElements(sapModel, groupName);

            for (int i = 0; i < objectNames.Length; i++)
            {
                EtabsObjectType objectType = (EtabsObjectType)objectTypeIds[i];
                if (!targetObjectTypes.Contains(objectType)) { continue; }  // Skip if not in target types

                objectTypeIdsListAfterFilter.Add(objectTypeIds[i]);
                objectNamesAfterFilter.Add(objectNames[i]);
            }

            return (objectTypeIdsListAfterFilter.ToArray(), objectNamesAfterFilter.ToArray());
        }
        #endregion

        #region ETABS Storey
        public static (Dictionary<string, double> storeyToElevationMap, Dictionary<double, string> elevationToStoreyMap) GetEtabsStoreys(cSapModel sapModel)
        {
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

            #region Map to Object
            Dictionary<string, double> storeyToElevationMap = new Dictionary<string, double>();
            Dictionary<double, string> elevationToStoreyMap = new Dictionary<double, string>();

            for (int i = 0; i < NumberStories; i++)
            {
                storeyToElevationMap.Add(storyNames[i], storyElevations[i]);
                elevationToStoreyMap.Add(Math.Round(storyElevations[i], 4), storyNames[i]);
            }
            #endregion

            return (storeyToElevationMap, elevationToStoreyMap);
        }

        public static string[] GetStoreyNames(cSapModel sapModel)
        {
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

            return storyNames;
        }

        public static string GetFirstStoreyName(cSapModel sapModel)
        {
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
            return storyNames[0];
            
        }
        #endregion

        #region Init
        public static void InitializeETABS(out ETABSv1.cOAPI etabsObject, out ETABSv1.cSapModel sapModel, bool setUnits = false)
        {
            etabsObject = null;
            sapModel = default(ETABSv1.cSapModel);

            try
            {
                etabsObject = (ETABSv1.cOAPI)Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");
                sapModel = etabsObject.SapModel;
                if (sapModel == null)
                {
                    throw new Exception("No active instance of ETABS found.");
                }

                if (setUnits)
                {
                    sapModel.SetPresentUnits_2(ETABSv1.eForce.kN, ETABSv1.eLength.m, ETABSv1.eTemperature.C);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Unable to attach to ETABS.\n" +
                    $"Check that active instance of API is set. See Tools -> Active Instance for API\n" +
                    $"{ex.Message}");
            }
        }
        #endregion

        #region Get Joints
        static (List<string> selectedJoints, List<double> Xs, List<double> Ys, List<double> Zs) GetSelectedJointsAsList(ETABSv1.cSapModel sapModel)
        {
            #region Get selected objects
            int numSel = 0;
            int[] objectType = new int[0];
            string[] objectName = new string[0];
            int ret = sapModel.SelectObj.GetSelected(ref numSel, ref objectType, ref objectName);
            if (ret != 0) { throw new Exception("Error getting selected joints"); }
            #endregion

            #region Get Point Coordinates
            List<string> selectedJoints = new List<string>();
            List<double> Xs = new List<double>();
            List<double> Ys = new List<double>();
            List<double> Zs = new List<double>();

            for (int i = 0; i < numSel; i++)
            {
                if (objectType[i] != 1) { continue; }
                selectedJoints.Add(objectName[i]);
                double x = 0;
                double y = 0;
                double z = 0;
                ret = sapModel.PointObj.GetCoordCartesian(objectName[i], ref x, ref y, ref z);
                if (ret != 0) { throw new Exception($"Error getting coordinate for joint {objectName[i]}"); }
                Xs.Add(Math.Round(x, 4));
                Ys.Add(Math.Round(y, 4));
                Zs.Add(Math.Round(z, 4));
            }
            #endregion

            if (selectedJoints.Count == 0) { throw new Exception($"No joints selected in ETABS"); }
            return (selectedJoints, Xs, Ys, Zs);
        }

        static (string[] selectedJoints, double[] Xs, double[] Ys, double[] Zs) GetSelectedJointsAsArray(ETABSv1.cSapModel sapModel)
        {
            (List<string> selectedJointsL, List<double> XsL, List<double> YsL, List<double> ZsL) = GetSelectedJointsAsList(sapModel);
            string[] selectedJoints = selectedJointsL.ToArray();
            double[] Xs = XsL.ToArray();
            double[] Ys = YsL.ToArray();
            double[] Zs = ZsL.ToArray();

            return (selectedJoints, Xs, Ys, Zs);
        }

        public static (string[] selectedJoints, double[] Xs, double[] Ys, double[] Zs) GetSortedJoints(ETABSv1.cSapModel sapModel, string sortType = "Z, X, Y")
        {
            (List<string> selectedJoints, List<double> Xs, List<double> Ys, List<double> Zs) = GetSelectedJointsAsList(sapModel);
            var jointObject = selectedJoints.Select((selectedJoint, i) => new { name = selectedJoint, x = Xs[i], y = Ys[i], z = Zs[i] });

            IEnumerable<(string name, double x, double y, double z)> sortedJoints;
            if (sortType == "Z, X, Y")
            {
                sortedJoints = jointObject
                .OrderBy(item => item.z)
                .ThenBy(item => item.x)
                .ThenBy(item => item.y)
                .Select(item => (item.name, item.x, item.y, item.z));
            }
            else if (sortType == "Z, Y, X")
            {
                sortedJoints = jointObject
                .OrderBy(item => item.z)
                .ThenBy(item => item.y)
                .ThenBy(item => item.x)
                .Select(item => (item.name, item.x, item.y, item.z));
            }
            else { throw new NotImplementedException($"Sort type \"{sortType}\" not implemented"); }


            string[] sortedJointArray = (string[])sortedJoints.Select(item => item.name).ToArray();
            double[] sortedXs = sortedJoints.Select(item => item.x).ToArray();
            double[] sortedYs = sortedJoints.Select(item => item.y).ToArray();
            double[] sortedZs = sortedJoints.Select(item => item.z).ToArray();

            return (sortedJointArray, sortedXs, sortedYs, sortedZs);
        }

        #endregion

        #region Break ETABS Table

        public static (string[,]tableData, string[]fieldKeysIncluded) GetEtabsTable2D(cSapModel sapModel, string tableName, string groupName = "All")
        {
            string[] fieldKeyList = null;
            string[] fieldsKeysIncluded = null;
            string[] tableData = null;
            int tableVersion = 0;
            int numberRecords = 0;

            int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                tableName,
                ref fieldKeyList,
                groupName,
                ref tableVersion,
                ref fieldsKeysIncluded,
                ref numberRecords,
                ref tableData
            );
            if (ret != 0) { throw new Exception("Error retrieving joint reaction data table from ETABS"); }

            string[,] tableData2d = BreakEtabsTableTo2D(tableData, fieldsKeysIncluded, numberRecords);
            return (tableData2d, fieldsKeysIncluded);
        }

        public static (Dictionary<string, string[]> tableDataDic, string[] fieldKeysIncluded, int numberRecords) GetEtabsTableDic(cSapModel sapModel, string tableName, string groupName = "All")
        {
            string[] fieldKeyList = null;
            string[] fieldsKeysIncluded = null;
            string[] tableData = null;
            int tableVersion = 0;
            int numberRecords = 0;

            int ret = sapModel.DatabaseTables.GetTableForDisplayArray(
                tableName,
                ref fieldKeyList,
                groupName,
                ref tableVersion,
                ref fieldsKeysIncluded,
                ref numberRecords,
                ref tableData
            );
            if (ret != 0) { throw new Exception("Error retrieving joint reaction data table from ETABS"); }

            Dictionary<string, string[]> tableDataDic = BreakEtabsTableToDictionary(tableData, fieldsKeysIncluded, numberRecords);
            return (tableDataDic, fieldsKeysIncluded, numberRecords);
        }

        public static string[,] BreakEtabsTableTo2D(string[] tableData, string[] fieldKeysIncluded, int numberRecords)
        {
            string[,] tableData2d = new string[numberRecords, fieldKeysIncluded.Length];
            int counter = 0;
            for (int rowNum = 0; rowNum < numberRecords; rowNum++)
            {
                for (int colNum = 0; colNum < fieldKeysIncluded.Length; colNum++)
                {
                    tableData2d[rowNum, colNum] = tableData[counter];
                    counter++;
                }
            }
            return tableData2d;
        }

        public static Dictionary<string,string[]> BreakEtabsTableToDictionary(string[] tableData, string[] fieldKeysIncluded, int numberRecords)
        {
            Dictionary<string, string[]> tableDataDic = new Dictionary<string, string[]>();
            
            #region Create Initial Empty Dictionary with Keys
            foreach (string fieldKey in fieldKeysIncluded)
            {
                tableDataDic.Add(fieldKey, new string[numberRecords]);
            }
            #endregion

            int counter = 0;
            for (int rowNum = 0; rowNum < numberRecords; rowNum++)
            {
                for (int colNum = 0; colNum < fieldKeysIncluded.Length; colNum++)
                {
                    string fieldKey = fieldKeysIncluded[colNum];
                    string[] dataArray = tableDataDic[fieldKey];
                    dataArray[rowNum] = tableData[counter];
                    counter++;
                }
            }
            return tableDataDic;
        }
        #endregion

        #region ETABS Group
        public static HashSet<string> GetExistingEtabsGroup(cSapModel sapModel, bool includeAll)
        {
            
            int numberNames = 0;
            string[] groupNames = new string[0];
            int ret = sapModel.GroupDef.GetNameList(ref numberNames, ref groupNames);
            if (ret != 0) { throw new Exception($"Unable to get group name list from ETABS"); }
            HashSet<string> groupNameSet = new HashSet<string>(groupNames);
            groupNameSet.Remove("All");
            return groupNameSet;
        }
        public static void CheckEtabsGroupExists(cSapModel sapModel, string[] groupNames)
        {
            HashSet<string> groupNameSet = GetExistingEtabsGroup(sapModel, false);
            List<string> undefinedGroups = new List<string>();
            foreach (string groupName in groupNames)
            {

                if (groupNameSet.Contains(groupName)) { continue; }
                undefinedGroups.Add(groupName);
            }

            if (undefinedGroups.Count > 0)
            {
                string undefinedGroupStr = string.Join(", ", undefinedGroups);
                throw new Exception($"The following ETABS groups do not exist: {undefinedGroupStr}");
            }
        }
        public static void CreateGroupIfNotExist(cSapModel sapModel, string[] groupNames)
        {
            HashSet<string> groupNameSet = GetExistingEtabsGroup(sapModel, false);
            List<string> undefinedGroups = new List<string>();
            foreach (string groupName in groupNames)
            {

                if (groupNameSet.Contains(groupName)) { continue; }
                undefinedGroups.Add(groupName);
                sapModel.GroupDef.SetGroup(groupName);
            }
        }
        #endregion

        #region Get UN from Pier Label
        public static Dictionary<string, List<string>> GetWallUNFromPierLabel(cSapModel sapModel, string[] labels, string targetSty = "")
        {
            #region Init Return Dictionary
            Dictionary<string, List<string>> pierLabeltoUnMap = new Dictionary<string, List<string>>();
            foreach (string pierLabel in labels)
            {
                if (string.IsNullOrEmpty(pierLabel)) { continue; }
                pierLabeltoUnMap.Add(pierLabel, new List<string>());
            }
            #endregion

            #region Get all Pier Labels
            HashSet<string> allPierLabels;
            {
                string[] allPierLabelsString = new string[0];
                int numberNames = 0;

                int ret = sapModel.PierLabel.GetNameList(ref numberNames, ref allPierLabelsString);
                allPierLabels = new HashSet<string>(allPierLabelsString);
            }
            #endregion

            #region Get all Walls
            string[] allAreaNames = new string[0];
            {
                int numberNames = 0;
                allAreaNames = new string[0];
                if (targetSty != "")
                {
                    int ret = sapModel.AreaObj.GetNameListOnStory(targetSty, ref numberNames, ref allAreaNames);
                }
                else
                {
                    int ret = sapModel.AreaObj.GetNameList(ref numberNames, ref allAreaNames);
                }
            }
            #endregion

            #region Add UN to target Pier Labels
            foreach (string areaUn in allAreaNames)
            {
                try
                {
                    string pierLabel = "";
                    int ret = sapModel.AreaObj.GetPier(areaUn, ref pierLabel);
                    if (pierLabeltoUnMap.ContainsKey(pierLabel))
                    {
                        pierLabeltoUnMap[pierLabel].Add(areaUn);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error getting pier label for area {areaUn}: {ex.Message}");
                }
            }
            #endregion

            return pierLabeltoUnMap;
        }
        #endregion
    }

    #region ETABS Objects
    public class GeneralEtabsObject
    {
        public string uniqueName;
        public int objectTypeInt;
        public EtabsObjectType objectType;
        public string labelName;
        public string status;
        public string storeyName;

        //public string objectTypeString;
        public GeneralEtabsObject(string uniqueName, string labelName, int objectTypeInt)
        {
            this.uniqueName = uniqueName;
            this.labelName = labelName;
            this.objectTypeInt = objectTypeInt;
            objectType = EtabsObjectTypeHelper.MapToObjectType(objectTypeInt);
        }
        // Should add an overrided class to get UN probably
        public string objectTypeString
        {
            get 
            { 
                switch (objectType)
                {
                    case EtabsObjectType.Point:
                        return "Joint";
                    case EtabsObjectType.Frame:
                        return "Frame";
                    case EtabsObjectType.Cable:
                        return "Cable";
                    case EtabsObjectType.Tendon:
                        return "Tendon";
                    case EtabsObjectType.Area:
                        return "Area";
                    case EtabsObjectType.Solid:
                        return "Solid";
                    case EtabsObjectType.Link:
                        return "Link";
                    default:
                        return "Unknown";
                }
            }
        }

            
    }

    #region Joints
    public class EtabsJoint : GeneralEtabsObject
    {
        public double x = double.NaN;
        public double y = double.NaN;
        public double z = double.NaN;
        
        public List<string> unstableDimension = new List<string>();

        public EtabsJoint(string uniqueName, string labelName) : base(uniqueName, labelName, 1)
        { 

        }
        public EtabsJoint(string uniqueName, string labelName, double x, double y, double z) : base(uniqueName, labelName, 1)
        {
            this.x = x;
            this.y = y;
            this.z = z;
        }

        #region Get UN/Labels
        public string GetAndSetUn(cSapModel sapModel)
        {
            int ret = sapModel.PointObj.GetNameFromLabel(labelName, storeyName, ref uniqueName);
            if (ret != 0) { throw new Exception($"Unable to get unique name for joint {labelName}, at storey {storeyName}"); }

            return uniqueName;
        }

        public (string,string) GetLabelAndStorey(cSapModel sapModel, Dictionary<double, string> elevationToStoreyMap)
        {
            if (uniqueName[0] == '~')
            {
                labelName = "NA";
                if (elevationToStoreyMap.ContainsKey(z / 1000))
                {
                    storeyName = elevationToStoreyMap[z / 1000];
                }
                else { storeyName = "Unknown"; }
            }
            else
            {
                int ret = sapModel.PointObj.GetLabelFromName(uniqueName, ref labelName, ref storeyName);
                if (ret != 0) { throw new Exception($"Unable to get label and storey for joint with unique name {uniqueName}"); }
            }
            return (labelName, storeyName);
        }
        #endregion

        #region Unstable Dimension
        public List<string> AddUnstableDimension(string dimension)
        {
            unstableDimension.Add(dimension);
            return unstableDimension;
        }

        public string GetAllUnstableDimensionsAsString()
        {
            string unstableDimString = string.Join(", ", unstableDimension);
            return unstableDimString;
        }
        #endregion

        #region Base Shear Reaction
        public Dictionary<string, double[]> baseReactions;
        public Dictionary<string, int> baseReactionRowNum;
        public void AddBaseReaction(string loadCase, int rowNum,
            double Fx, double Fy, double Fz, double Mx, double My, double Mz)
        {
            #region Checks
            if (baseReactions == null) { 
                baseReactions = new Dictionary<string, double[]>(); 
            }
            if (baseReactionRowNum == null) { baseReactionRowNum = new Dictionary<string, int>(); }
            if (baseReactions.ContainsKey(loadCase)) { throw new Exception($"Load case \"{loadCase}\" already exist for joint with unique name \"{uniqueName}\""); }
            if (baseReactionRowNum.ContainsKey(loadCase)) { throw new Exception($"Load case \"{loadCase}\" already exist for joint with unique name \"{uniqueName}\""); }
            #endregion

            double[] reactions = new double[9];
            reactions[0] = Fx;
            reactions[1] = Fy;
            reactions[2] = Fz;
            reactions[3] = Mx;
            reactions[4] = My;
            reactions[5] = Mz;
            baseReactions.Add(loadCase, reactions);
            baseReactionRowNum.Add(loadCase, rowNum);
        }
        /// <summary>
        ///     Assumes all joints are at base level, no consideration of elevation
        /// </summary>
        public double[] GetMomentAboutPoint(cSapModel sapModel, string loadCase, double xOrigin, double yOrigin)
        {
            //throw new NotImplementedException("Not working, kept for reference");
            if (double.IsNaN(this.x)) { GetCoordinates(sapModel); }

            if (!baseReactions.ContainsKey(loadCase)) { throw new Exception($"Base reaction for joint with unqiue name \"{uniqueName}\" and load case \"{loadCase}\" not initialised, unable to calculate base moment."); }
            double[] reactions = baseReactions[loadCase];

            double dx = this.x - xOrigin;
            double dy = this.y - yOrigin;
            if (double.IsNaN(dx) || double.IsNaN(dy)) { throw new Exception($"Undefined origin for joint with UN {uniqueName}"); }

            reactions[6] = dy * baseReactions[loadCase][2]; // Mx about origin
            reactions[7] = -dx * baseReactions[loadCase][2]; // My about origin
            reactions[8] = dx * baseReactions[loadCase][1] - dy * baseReactions[loadCase][0]; // Mz about origin

            return reactions;
        }
        #endregion

        #region Get Coordinates
        public double[] GetCoordinates(cSapModel sapModel)
        {
            int ret = sapModel.PointObj.GetCoordCartesian(uniqueName, ref x, ref y, ref z);
            if (ret != 0) { throw new Exception($"Error getting coordinate for joint with unique name {uniqueName}"); }
            double[] coordinates = new double[3] { x, y, z };
            return coordinates;
        }
        #endregion
    }
    #endregion

    #region Frame
    public class EtabsFrame: GeneralEtabsObject
    {
        public string[] jointUN = new string[2];
        public eFrameDesignOrientation frameType;
        public EtabsFrame(string uniqueName, string labelName, cSapModel sapModel = null) : base(uniqueName, labelName, 2)
        {
            // If sapModel provided, classify immediately
            if (sapModel != null) { Classify(sapModel); }
        }
        public eFrameDesignOrientation Classify(cSapModel sapModel)
        {
            frameType = eFrameDesignOrientation.Null;
            int ret = sapModel.FrameObj.GetDesignOrientation(uniqueName, ref frameType);

            if (ret != 0)
            {
                throw new Exception($"Error: Frame with unique name {uniqueName} not found or API Failure");
            }
            return frameType;
        }
        public string[] GetBaseJoints(cSapModel sapModel)
        {
            throw new NotImplementedException();
        }


        #region Get Details

        #region Joints
        List<EtabsJoint> joints;
        public List<EtabsJoint> GetJoints(cSapModel sapModel)
        {
            if (joints != null) { return joints; }
            joints = new List<EtabsJoint>();
            string point1 = "";
            string point2 = "";

            int ret = sapModel.FrameObj.GetPoints(uniqueName, ref point1, ref point2);
            if (ret != 0)
            {
                throw new Exception($"Error: Area with unique name {uniqueName} not found or API Failure");
            }

            joints.Add(new EtabsJoint(point1, ""));
            joints.Add(new EtabsJoint(point2, ""));
            return joints;
        }
        public double[] coordinates = new double[6]; // x1, y1, z1, x2, y2, z2
        public double[] GetCoordinates(cSapModel sapModel)
        {
            if (joints == null) { GetJoints(sapModel); }
            if (joints.Count != 2) { throw new Exception($"Frame with unique name {uniqueName} has {joints.Count} joints. Unexpected result"); }

            int i = 0;
            foreach (EtabsJoint joint in joints)
            {
                double[] jointCoord = joint.GetCoordinates(sapModel);
                foreach (double coord in jointCoord)
                {
                    coordinates[i] = coord;
                    i += 1;
                }
            }
            return coordinates;
        }
        #endregion
        public string GetLabel(cSapModel sapModel)
        {
            sapModel.FrameObj.GetLabelFromName(uniqueName, ref labelName, ref storeyName);
            return labelName;
        }

        public string sectionName = "";
        public string GetSection(cSapModel sapModel)
        {
            string autoSelect = "";
            sapModel.FrameObj.GetSection(uniqueName, ref sectionName, ref autoSelect);
            return sectionName;
        }

        #endregion

        #region Sets
        public void SetUniqueName(cSapModel sapModel, string newName, bool setErrAsStatus = false)
        {
            int ret = sapModel.FrameObj.ChangeName(uniqueName, newName);
            if (ret != 0)
            {
                throw new Exception($"Error changing name of frame from {uniqueName} to {newName}");
            }
        }

        public void SetSection(cSapModel sapModel, string section, bool setErrAsStatus = false)
        {
            int ret = sapModel.FrameObj.SetSection(uniqueName, section);
            if (ret != 0)
            {
                throw new Exception($"Error changing section of frame to {section}");
            }
        }
        #endregion
    }

    #endregion

    #region Area

    public class EtabsArea : GeneralEtabsObject
    {
        public string[] jointUN = new string[4];
        eAreaDesignOrientation areaType;
        public EtabsArea(string uniqueName, string labelName, cSapModel sapModel = null) : base(uniqueName, labelName, 5)
        {
            // If sapModel provided, classify immediately
            if (sapModel != null) { Classify(sapModel); }
        }
        public eAreaDesignOrientation Classify(cSapModel sapModel)
        {
            areaType = eAreaDesignOrientation.Null;
            int ret = sapModel.AreaObj.GetDesignOrientation(uniqueName, ref areaType);

            if (ret != 0)
            {
                throw new Exception($"Error: Area with unique name {uniqueName} not found or API Failure");
            }
            return areaType;
        }

        public string[] GetBaseJoints(cSapModel sapModel)
        {
            throw new NotImplementedException();
        }

        #region Joints
        List<EtabsJoint> joints;
        public List<EtabsJoint> GetJoints(cSapModel sapModel)
        {
            if (joints != null) { return joints; }
            joints = new List<EtabsJoint>();
            int numPoints = -1;
            string[] pointNames = new string[0];
            int ret = sapModel.AreaObj.GetPoints(uniqueName, ref numPoints, ref pointNames);
            if (ret != 0)
            {
                throw new Exception($"Error: Area with unique name {uniqueName} not found or API Failure");
            }

            foreach (string pointName in pointNames)
            {
                EtabsJoint jointObj = new EtabsJoint(pointName, "");
                joints.Add(jointObj);
            }
            return joints;
        } 
        #endregion
    }
    #endregion

    #region ETABS Object Type
    static public class EtabsObjectTypeHelper
    {
        static public EtabsObjectType MapToObjectType(int objectTypeInt, string uniqueName = "", string labelName = "")
        {
            if (objectTypeInt < 1 || objectTypeInt > 7)
            {
                throw new Exception($"Object type {objectTypeInt} not recognised for object with UN: {uniqueName}, Label Name: {labelName}");
            }
            EtabsObjectType objectType = (EtabsObjectType)objectTypeInt;
            return objectType;
        }
        static public GeneralEtabsObject ClassifyEtabsObject(string uniqueName, int objectTypeInt, cSapModel sapModel = null)
        {
            switch (objectTypeInt)
            {

                case 1: //Point
                    {
                        return new EtabsJoint(uniqueName, "");
                    }
                case 2: //Frame
                    {
                        return new EtabsFrame(uniqueName, "", sapModel);
                    }

                case 3: //Cable
                    {
                        return new GeneralEtabsObject(uniqueName, "", objectTypeInt);
                    }

                case 4: //Tendon
                    {
                        return new GeneralEtabsObject(uniqueName, "", objectTypeInt);
                    }

                case 5: //Area
                    {
                        return new EtabsArea(uniqueName, "", sapModel);
                    }

                case 6: //Solid
                    {
                        return new GeneralEtabsObject(uniqueName, "", objectTypeInt);
                    }

                case 7: //Link
                    {
                        return new GeneralEtabsObject(uniqueName, "", objectTypeInt);
                    }
                default:
                    throw new Exception($"Object type {objectTypeInt} not recognised");
            }
        }
        static public List<GeneralEtabsObject> GetVerticalElements(cSapModel sapModel, string groupName)
        {
            (int[] objectTypeIds, string[] objectNames) = EtabsFunctions.GetGroupElementofType(sapModel, groupName, new HashSet<EtabsObjectType> { EtabsObjectType.Area, EtabsObjectType.Frame });
            //Dictionary<string, GeneralEtabsObject> colAndWallObjects = new Dictionary<string, GeneralEtabsObject>();
            List<GeneralEtabsObject> colAndWallObjects = new List<GeneralEtabsObject>();
            for (int i = 0; i < objectTypeIds.Length; i++)
            {
                GeneralEtabsObject obj = ClassifyEtabsObject(objectNames[i], objectTypeIds[i], sapModel);
                if (obj.objectType == EtabsObjectType.Area)
                {
                    EtabsArea areaObj = (EtabsArea)obj;
                    eAreaDesignOrientation areaType = areaObj.Classify(sapModel);
                    if (areaType != eAreaDesignOrientation.Wall)
                    {
                        continue;
                    }
                }
                else if (obj.objectType == EtabsObjectType.Frame)
                {
                    EtabsFrame frameObj = (EtabsFrame)obj;
                    eFrameDesignOrientation frameType = frameObj.Classify(sapModel);
                    if (frameType != eFrameDesignOrientation.Column && frameType != eFrameDesignOrientation.Brace)
                    {
                        continue;
                    }
                }
                else
                {
                    continue;
                }
                colAndWallObjects.Add(obj);
            }
            return colAndWallObjects;
        }
        static public Dictionary<string, GeneralEtabsObject> GetUniqueJointsFromElements(cSapModel sapModel, IEnumerable<GeneralEtabsObject> elements, bool throwWarningForUndefinedElements = true)
        {
            Dictionary<string, GeneralEtabsObject> jointsDict = new Dictionary<string, GeneralEtabsObject>();
            foreach (var element in elements)
            {
                //List<EtabsJoint> joints = new List<EtabsJoint>();
                switch (element.objectType)
                {
                    case EtabsObjectType.Point:
                        {
                            EtabsJoint jointObj = (EtabsJoint)element;

                            if (!jointsDict.ContainsKey(jointObj.uniqueName))
                            {
                                jointsDict.Add(jointObj.uniqueName, jointObj);
                            }
                            break;
                        }
                    case EtabsObjectType.Frame:
                        {
                            EtabsFrame frameObj = (EtabsFrame)element;

                            foreach (EtabsJoint jointObj in frameObj.GetJoints(sapModel))
                            {
                                if (!jointsDict.ContainsKey(jointObj.uniqueName))
                                {
                                    jointsDict.Add(jointObj.uniqueName, jointObj);
                                }
                            }
                            break;
                        }
                    case EtabsObjectType.Area:
                        {
                            EtabsArea areaObj = (EtabsArea)element;

                            foreach (EtabsJoint jointObj in areaObj.GetJoints(sapModel))
                            {
                                if (!jointsDict.ContainsKey(jointObj.uniqueName))
                                {
                                    jointsDict.Add(jointObj.uniqueName, jointObj);
                                }
                            }
                            break;
                        }
                    default:
                        {
                            if (throwWarningForUndefinedElements)
                            {
                                throw new Exception($"Object type {element.objectType} not supported for getting joints");
                            }
                            break;
                        }
                }
            }

            return jointsDict;
        }
        static public EtabsFrame[] GetSpecificFrameType(string[] frameUns, cSapModel sapModel, HashSet<eFrameDesignOrientation> targetFrameTypes)
        {
            List<EtabsFrame> targetFrameObjects = new List<EtabsFrame>();
            foreach (string frameUn in frameUns)
            {
                EtabsFrame frameObj = new EtabsFrame(frameUn, "", sapModel);
                bool isTarget = false;
                if (targetFrameTypes.Count == 0) { isTarget = true; } // No target, get all
                else if (targetFrameTypes.Contains(frameObj.frameType)) { isTarget = true; } // Is target type
                if (isTarget) { targetFrameObjects.Add(frameObj); }
            }
            return targetFrameObjects.ToArray();
        }

        static public void GetFrameDetails(cSapModel sapModel, ref EtabsFrame frameObj, bool getLabel, bool getCoord, bool getSection)
        {
            if (getLabel)
            {
                frameObj.GetLabel(sapModel);
            }

            if (getCoord)
            {
                frameObj.GetCoordinates(sapModel);
            }

            if (getSection)
            {
                frameObj.GetSection(sapModel);
            } 
        }

    }
    public enum EtabsObjectType
    {
        Point = 1, 
        Frame = 2, 
        Cable = 3, 
        Tendon = 4,
        Area = 5, 
        Solid = 6, 
        Link = 7 
    }

    #region References
    //eAreaDesignOrientation
    //Wall 1  
    //Floor 2  
    //Ramp_DO_NOT_USE 3  
    //Null 4  
    //Other 5 

    //eFrameDesignOrientation
    //Column 1  
    //Beam 2  
    //Brace 3  
    //Null 4  
    //Other 5 
    #endregion
    #endregion

    #endregion


}

