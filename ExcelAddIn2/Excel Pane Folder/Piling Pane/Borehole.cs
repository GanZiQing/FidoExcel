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
    class Borehole
    {
        public string name;
        public double reducedLevel;
        public double index;
        public double[] depth;
        public double[] sptValue;
        public double[] excelLoc;
        public int[] isRock;
        public string[] rockType;


        public Borehole(Range readRange, HashSet<string> rockTypes, HashSet<string> notRockTypes)
        {
            ReadRange(readRange, rockTypes, notRockTypes);
        }

        static void ThrowExceptionBox(string msg)
        {
            Console.WriteLine($"{msg}");
            throw new Exception(msg);
        }

        private void ReadRange(Range readRange, HashSet<string> rockTypes, HashSet<string> notRockTypes)
        {
            Worksheet thisSheet = readRange.Worksheet;
            int startRowNum = readRange.Row;
            int startColNum = readRange.Column;

            // Save excel location info
            excelLoc = new double[2];
            excelLoc[0] = startRowNum;
            excelLoc[1] = startColNum;

            #region Loop Through Range
            // Row 1 & 2 is fixed
            name = thisSheet.Cells[startRowNum, startColNum].Value2.ToString();
            reducedLevel = ReadDoubleFromCell(thisSheet.Cells[startRowNum + 1, startColNum + 1]);
            index = ReadDoubleFromCell(thisSheet.Cells[startRowNum + 2, startColNum + 1]);
            // Row 4 onwards is data
            int numInputs = 0;
            int maxNumCol = 3;
            List<double> depthList = new List<double>();
            List<double> sptValueList = new List<double>();
            List<int> isRockList = new List<int>();
            List<string> rockTypeList = new List<string>();

            for (int locColNum = 0; locColNum < maxNumCol; locColNum += 1)
            {
                int locRowNum = 0;
                while (numInputs == 0 || (numInputs != 0 && locRowNum < numInputs))
                {
                    int rowNum = locRowNum + startRowNum + 5;
                    int colNum = locColNum + startColNum;

                    Range cell = thisSheet.Cells[rowNum, colNum];
                    // Break if empty row
                    if (cell.Value2 == null && numInputs == 0)
                    {
                        numInputs = locRowNum;
                        break;
                    }

                    double cellValue;
                    switch (locColNum)
                    {
                        case 0:
                            {
                                List<double> thisList = depthList;
                                cellValue = ReadDoubleFromCell(cell);
                                if (double.IsNaN(cellValue))
                                {
                                    cell.Select();
                                    //MessageBox.Show($"Warning, NaN value read at {cell.get_Address(RowAbsolute: false, ColumnAbsolute: false)}, please check result.");
                                    //Beaver.LogError($"Warning, NaN value read at {cell.get_Address(RowAbsolute: false, ColumnAbsolute: false)}, please check result.");
                                }
                                thisList.Add(cellValue);
                            }
                            break;
                        case 1:
                            {
                                List<double> thisList = sptValueList;
                                cellValue = ReadDoubleFromCell(cell);
                                if (double.IsNaN(cellValue))
                                {
                                    //MessageBox.Show($"Warning, NaN value read at {cell.get_Address(RowAbsolute: false, ColumnAbsolute: false)}, please check result.");
                                    //Beaver.LogError($"Warning, NaN value read at {cell.get_Address(RowAbsolute: false, ColumnAbsolute: false)}, please check result.");
                                }
                                thisList.Add(cellValue);
                            }
                            break;
                        case 2:
                            {
                                if (cell.Value2 != null)
                                {
                                    string inputRockType = cell.Value2.ToString();
                                    inputRockType = inputRockType.Trim();
                                    rockTypeList.Add(inputRockType);

                                    if (rockTypes.Contains(inputRockType))
                                    {    
                                        isRockList.Add(1); // Is rock
                                    }
                                    else if (notRockTypes.Contains(inputRockType))
                                    {   
                                        isRockList.Add(0); // Is overwrite rock
                                    }
                                    else
                                    {
                                        isRockList.Add(-1); // Is not rock
                                    }
                                }
                                else
                                {
                                    rockTypeList.Add("");
                                    isRockList.Add(0);
                                }
                                
                            }
                            break;
                        default:
                            MessageBox.Show("Column number exceeded");
                            return;
                    }
                    locRowNum += 1;
                }
            }
            #endregion

            depth = depthList.ToArray<double>();
            sptValue = sptValueList.ToArray<double>();
            isRock = isRockList.ToArray<int>();
            rockType = rockTypeList.ToArray<string>();
        }
    }
}
