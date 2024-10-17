using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public class ExcelTableCol
    {
        public int relativeColNum;
        public string name;
        public Range range;

        public ExcelTableCol(string name, int relativeColNum = 0)
        {
            this.name = name;
            if (relativeColNum != 0)
            {
                this.relativeColNum = relativeColNum;
            }
        }

        #region Convert to Array
        public string[] stringArray;
        public string[] ConvertRangeToStringArray()
        {
            if (range == null)
            {
                throw new Exception($"Unable to convert range to array.\nRange not defined for {name}");
            }
            stringArray = new string[range.Cells.Count];
            int index = 0;
            foreach (Range cell in range.Cells)
            {
                if (cell.Value2 == null)
                {
                    stringArray[index] = "";
                }
                else
                {
                    stringArray[index] = cell.Value2.ToString();
                }
                index++;
            }
            return stringArray;
        }

        public bool[] boolArray;
        public bool[] ConvertRangeToBoolArray()
        {
            if (range == null)
            {
                throw new Exception($"Unable to convert range to array.\nRange not defined for {name}");
            }
            HashSet<string> trueValues = new HashSet<string> { "Yes", "yes", "YES", "True", "TRUE" };
            boolArray = new bool[range.Cells.Count];
            int index = 0;
            foreach (Range cell in range.Cells)
            {
                if (cell.Value2 == null)
                {
                    boolArray[index] = false;
                }
                else if (cell.Value2 is bool)
                {
                    boolArray[index] = (bool)cell.Value2;
                }
                else
                {
                    string cellString = cell.Value2.ToString();
                    cellString = cellString.Trim();
                    if (trueValues.Contains(cellString))
                    {
                        boolArray[index] = true;
                    }
                    else
                    {
                        boolArray[index] = false;
                    }
                }

                index++;
            }
            return boolArray;
        }

        public int?[] intArray;
        public int?[] CreateNewIntArray()
        {
            if (range == null)
            {
                throw new Exception($"Unable to convert range to array.\nRange not defined for {name}");
            }
            intArray = new int?[range.Cells.Count];
            return intArray;
        }
        #endregion

        #region Write to Excel
        public void WriteIntToExcel()
        {
            if (intArray == null)
            {
                throw new Exception("Int array not initialised");
            }
            if (intArray.Length > range.Cells.Count)
            {
                throw new Exception("Int array larger than write range");
            }

            int index = 0;
            foreach (Range cell in range.Cells)
            {
                if (intArray == null)
                {
                    continue;
                }
                cell.Value2 = intArray[index];
                index++;
            }
        }
        #endregion
    }
    public class ExcelTable

    {
        #region Initialisation
        Range tableRange;
        Dictionary<int, ExcelTableCol> positionToColumnDict;
        Dictionary<string, ExcelTableCol> nameToColumnDict;
        string name;
        public ExcelTable(Range tableRange, string name)
        {
            this.name = name;
            this.tableRange = tableRange;
            positionToColumnDict = new Dictionary<int, ExcelTableCol>();
            nameToColumnDict = new Dictionary<string, ExcelTableCol>();
        }
        #endregion

        #region Adding Columns
        public ExcelTableCol AddColumn(int columnPosition, string columnName)
        {
            ExcelTableCol thisEntry = new ExcelTableCol(columnName, columnPosition);
            positionToColumnDict.Add(columnPosition, thisEntry);
            nameToColumnDict.Add(columnName, thisEntry);
            return thisEntry;
        }
        #endregion

        #region Getting Column
        public ExcelTableCol GetColumnFromPosition(int colPosition)
        {
            return positionToColumnDict[colPosition];
        }

        public Range GetRangeFromPosition(int colPosition)
        {
            return positionToColumnDict[colPosition].range;
        }

        public ExcelTableCol GetColumnFromName(string name)
        {
            return nameToColumnDict[name];
        }

        public Range GetRangeFromName(string name)
        {
            return nameToColumnDict[name].range;
        }
        #endregion

        #region Reading Columns from Table
        public void ReadRangeFromTable()
        {
            if (positionToColumnDict.Keys.Count == 0)
            {
                throw new Exception($"Table Assignment for {name} is empty");
            }

            foreach (KeyValuePair<int, ExcelTableCol> entry in positionToColumnDict)
            {
                int relColNum = entry.Key;
                ExcelTableCol thisColumn = entry.Value;

                Range sourceCol = tableRange.Columns[relColNum];
                thisColumn.range = sourceCol;
            }
        }
        #endregion
    }
}
