using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;

namespace ExcelAddIn2.Piling
{
    class BoreholeAGS
    {
        public string name;
        public double reducedLevel;
        public double[] depth;
        public double[] sptValue;
        public string[] soilType;
        private string[] soilDescription;
        public double terminationDepth = 0;
        private double rockStartDepth;
        private double spt100StartDepth;


        private List<double> depthList;
        private List<string> soilTypeList;
        private List<string> soilDescriptionList;
        private List<double> sptValueList;
        private bool convertedToArray;


        public BoreholeAGS(bool ags = false)
        {
            //ReadRange(readRange);
            if (ags)
            {
                convertedToArray = false;
                depthList = new List<double>();
                soilTypeList = new List<string>();
                soilDescriptionList = new List<string>();
                sptValueList = new List<double>();
            }
        }

        static void ThrowExceptionBox(string msg)
        {
            Console.WriteLine($"{msg}");
            throw new Exception(msg);
        }

        public void AddBhSoilTypeToList(string inputTopDepthString, string inputBotDepthString,string inputSoilType, string inputSoilDescription)
        {
            double inputDepth = double.Parse(inputTopDepthString);
            if (depthList.Count > 0 && inputDepth < depthList.Last())
            {
                //MessageBox.Show($"Warning, borehole depth for {name} is not in sequence?\n" +
                //    $"Depth = {inputDepth}, Soil Type = {inputSoilType}\n" +
                //    $"Check input file.");
                Beaver.LogError($"Warning, unexpected behaviour encountered in soil type info. Borehole depth for {name} is not in sequence?\n" +
                    $"Depth = {inputDepth}, Soil Type = {inputSoilType}\n" +
                    $"Check input file at **GEOL.");
            }
            depthList.Add(inputDepth);
            soilTypeList.Add(inputSoilType);
            soilDescriptionList.Add(inputSoilDescription);

            double botDepth = double.Parse(inputBotDepthString);
            if (terminationDepth < botDepth)
            {
                terminationDepth = botDepth;
            }
        }

        public void AddSPT(string inputDepthString, string inputSptString)
        {
            double inputDepth = double.Parse(inputDepthString);
            double inputSpt = double.Parse(inputSptString);
            if (sptValueList.Count == 0)
            {
                for (int i = 0; i < depthList.Count; i++)
                {
                    sptValueList.Add(double.NaN);
                }
            }

            int insertionIndex = depthList.IndexOf(inputDepth);
            if (insertionIndex != -1)
            {
                sptValueList[insertionIndex] = inputSpt;
            }
            else if (inputDepth > depthList.Last())
            {
                depthList.Add( inputDepth);
                soilTypeList.Add("");
                soilDescriptionList.Add("");
                sptValueList.Add(inputSpt);
            }
            else
            {
                insertionIndex = depthList.FindIndex(depth => depth > inputDepth);
                depthList.Insert(insertionIndex, inputDepth);
                soilTypeList.Insert(insertionIndex, "");
                soilDescriptionList.Insert(insertionIndex, "");
                sptValueList.Insert(insertionIndex, inputSpt);
            }
        }

        public void AddBhRl(string bhRL)
        {
            bool isDouble = double.TryParse(bhRL, out reducedLevel);
            if (!isDouble)
            {
                Beaver.LogError($"Unable to parse reduce level {bhRL} into double for {name}");
            }
        }

        public void ProcessAndConvertToArray(HashSet<string> rockTypes, HashSet<string> spt100Types, bool skipEmptySPT = false, bool fillSoilType = false, bool defaultSPT = false, bool compressOutput = false)
        {
            #region Copy soil type to soil type to bottom if empty
            if (fillSoilType)
            {
                for (int i = 0; i < soilTypeList.Count-1; i++)
                {
                    if (soilTypeList[i + 1] == "")
                    {
                        soilTypeList[i + 1] = soilTypeList[i];
                    }
                    if (soilDescriptionList[i + 1] == "")
                    {
                        soilDescriptionList[i + 1] = soilDescriptionList[i];
                    }
                }
            }
            #endregion

            #region Fill default SPT Values at top and check rock/spt
            if (defaultSPT)
            {
                if (double.IsNaN(sptValueList[0]))
                {
                    sptValueList[0] = 0;
                }

                for (int i = 0; i < depthList.Count; i++)
                {
                    if (spt100Types.Contains(soilTypeList[i]))
                    {
                        sptValueList[i] = 100;
                    }
                    else if(rockTypes.Contains(soilTypeList[i]))
                    {
                        sptValueList[i] = 100;
                    }
                }

                if (double.IsNaN(sptValueList.Last()))
                {
                    sptValueList[sptValueList.Count - 1] = -1;
                    Beaver.LogError($"Final SPT value for {name} is empty, check output");
                }
            }

            #endregion

            #region Remove empty SPT rows
            if (skipEmptySPT)
            {
                int numRows = soilTypeList.Count;
                depth = new double[numRows];
                soilType = new string[numRows];
                soilDescription = new string[numRows];
                sptValue = new double[numRows];

                int newRowNum = 0;
                for (int i = 0; i < numRows; i++) // Do not remove last row
                {
                    if (double.IsNaN(sptValueList[i]))
                    {
                        continue;
                    }
                    else
                    {
                        depth[newRowNum] = depthList[i];
                        soilType[newRowNum] = soilTypeList[i];
                        soilDescription[newRowNum] = soilDescriptionList[i];
                        sptValue[newRowNum] = sptValueList[i];
                        newRowNum++;
                    }
                }
                Array.Resize(ref depth, newRowNum);
                Array.Resize(ref soilType, newRowNum);
                Array.Resize(ref soilDescription, newRowNum);
                Array.Resize(ref sptValue, newRowNum);
            }
            else
            {
                depth = depthList.ToArray();
                soilType = soilTypeList.ToArray();
                soilDescription = soilDescriptionList.ToArray();
                sptValue = sptValueList.ToArray();
            }
            #endregion

            #region Compress output - merge duplicate rows
            if (compressOutput)
            {
                int newRowNum = 0;
                for (int rowNum = 1; rowNum < depth.Length; rowNum++)
                {
                    if (soilType[newRowNum] != soilType[rowNum] || sptValue[newRowNum] != sptValue[rowNum])
                    {
                        newRowNum++;
                        depth[newRowNum] = depth[rowNum];
                        soilType[newRowNum] = soilType[rowNum];
                        soilDescription[newRowNum] = soilDescription[rowNum];
                        sptValue[newRowNum] = sptValue[rowNum];
                    }
                }
                Array.Resize(ref depth, newRowNum + 1);
                Array.Resize(ref soilType, newRowNum + 1);
                Array.Resize(ref soilDescription, newRowNum + 1);
                Array.Resize(ref sptValue, newRowNum + 1);
            }
            #endregion

            #region Find Rock Start
            rockStartDepth = double.NaN;
            for (int i = depth.Length - 1; i >= 1 ; i--)
            {
                if (soilType[i] == "")
                {
                    continue;
                }
                else if (rockTypes.Contains(soilType[i]))
                {
                    rockStartDepth = depth[i];
                }
                else
                {
                    break;
                }
            }
            #endregion

            #region Find SPT100 Start
            spt100StartDepth = double.NaN;
            for (int i = depth.Length - 1; i >= 1; i--)
            {
                if (double.IsNaN(sptValue[i]))
                {
                    continue;
                }
                else if (sptValue[i] == 100)
                {
                    spt100StartDepth = depth[i];
                }
                else
                {
                    break;
                }
            }
            #endregion

            depthList = null;
            soilTypeList = null;
            soilDescriptionList = null;
            sptValueList = null;
            convertedToArray = true;
        }

        public void WriteBHToExcel(Range selRange, HashSet<string> rockTypes, HashSet<string> spt100Types, bool skipEmptySPT = false, bool fillSoilType = false, bool writeDescription = false, bool defaultSPT = false, bool compressOutput = false)
        {
            if (!convertedToArray)
            {
                ProcessAndConvertToArray(rockTypes, spt100Types, skipEmptySPT, fillSoilType, defaultSPT, compressOutput);
            }

            selRange.Cells[1, 1].Value2 = name;
            selRange.Cells[2, 1].Value2 = "RL:";
            selRange.Cells[2, 2].Value2 = reducedLevel;
            selRange.Cells[3, 1].Value2 = "Index";
            selRange.Cells[5, 1].Value2 = "Depth";
            selRange.Cells[5, 2].Value2 = "SPT";
            selRange.Cells[5, 3].Value2 = "Soil Type";

            selRange.Cells[4, 1].Value2 = "SPT100/Rock Start:";

            if (!double.IsNaN(spt100StartDepth))
            {
                selRange.Cells[4, 2].Value2 = spt100StartDepth;
            }
            if (!double.IsNaN(rockStartDepth))
            {
                selRange.Cells[4, 3].Value2 = rockStartDepth;
            }

            

            Range dataRange = selRange.Offset[5, 0];
            Range writeRange = null;
            if (writeDescription)
            {
                selRange.Cells[5, 4].Value2 = "Soil Description";
                writeRange = WriteArrayToExcelRange(dataRange, 0, 0, depth, sptValue, soilType, soilDescription);
                
            }
            else
            {
                writeRange = WriteArrayToExcelRange(dataRange, 0, 0, depth, sptValue, soilType);
            }
            #region Write Termination Depth
            Range terminationCell = writeRange.Cells[writeRange.Rows.Count + 1, 1];
            terminationCell.Value2 = terminationDepth;
            #endregion
        }

        private Range WriteArrayToExcelRange(Range thisRange, int rowOff, int colOff, params Array[] arrays)
        {
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
                    object input = arrays[col].GetValue(row);
                    if (input is double && double.IsNaN((double)input))
                    {
                        input = "";
                    }
                    dataArray[row, col] = input;
                    //dataArray[row, col] = arrays[col].GetValue(row);
                }
            }

            // Write to Excel
            //thisRange.Application.ScreenUpdating = false;
            Range startCell = thisRange.Offset[rowOff, colOff];
            Range endCell = startCell.Offset[numRow - 1, numCol - 1];
            Range writeRange = thisRange.Worksheet.Range[startCell, endCell];
            writeRange.ClearContents();
            writeRange.Value2 = dataArray;
            //thisRange.Application.ScreenUpdating = true;
            return writeRange;
        }
    }
}
