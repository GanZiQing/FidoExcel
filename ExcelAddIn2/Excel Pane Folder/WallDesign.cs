using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Microsoft.Office.Interop.Excel;
using static ExcelAddIn2.CommonUtilities;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class WallDesign : UserControl
    {
        #region Init
        Dictionary<string, AttributeTextBox> TextBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();
        public WallDesign()
        {
            InitializeComponent();
            CreateAttributes();
        }

        private void CreateAttributes()
        {
            RangeTextBox rangeTB = new RangeTextBox("rebarTable_WD", dispRebarTable, setRebarTable, "range");
            TextBoxAttributeDic.Add(rangeTB.attName, rangeTB);

            rangeTB = new RangeTextBox("storeyTable_WD", dispStoreyTable, setStoreyTable, "range");
            TextBoxAttributeDic.Add(rangeTB.attName, rangeTB);

            rangeTB = new RangeTextBox("pierLabelRange_WD", dispPierLabelRange, setPierLabelRange, "column");
            TextBoxAttributeDic.Add(rangeTB.attName, rangeTB);

            rangeTB = new RangeTextBox("matchStoreyRange_WD", dispMatchStoreyCol, setMatchStoreyCol, "cell");
            TextBoxAttributeDic.Add(rangeTB.attName, rangeTB);

            rangeTB = new RangeTextBox("outputRange_WD", dispOutputCol, setOutputCol, "cell");
            TextBoxAttributeDic.Add(rangeTB.attName, rangeTB);

            rangeTB = new RangeTextBox("statusRange_WD", dispStatusCol, setStatusCol, "cell");
            TextBoxAttributeDic.Add(rangeTB.attName, rangeTB);
        }

        #endregion

        #region Match Rebars
        StoreyTracker storeyTracker; 
        private void matchWallRebar_Click(object sender, EventArgs e)
        {
            try
            {
                #region Storey
                Range storeyTable = ((RangeTextBox)TextBoxAttributeDic["storeyTable_WD"]).GetRangeFromFullAddress();
                storeyTracker = new StoreyTracker(storeyTable);
                #endregion


                #region Rebar Table
                ProcessRebarTable();
                #endregion

                #region Match Values
                MatchRebars();
                #endregion

                #region Release Dictionaries
                rebarDic = null;
                storeyTracker = null;
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #region Process Rebar Table
        Dictionary<string, WallRebar> rebarDic;
        private void ProcessRebarTable()
        {
            Range rebarTableRange = ((RangeTextBox)TextBoxAttributeDic["rebarTable_WD"]).GetRangeFromFullAddress();
            rebarDic = new Dictionary<string, WallRebar>();
            foreach (Range row in rebarTableRange.Rows)
            {
                string name = row.Cells[1].Text;
                if (!rebarDic.ContainsKey(name)) { rebarDic[name] = new WallRebar(name, storeyTracker); }
                WallRebar wallRebar = rebarDic[name];
                wallRebar.AddRow(row);
            }

            foreach (WallRebar wallRebar in rebarDic.Values)
            {
                wallRebar.SortStories();
            }
        }

        #endregion

        private void MatchRebars()
        {
            #region Get Ranges
            Range pierLableRange = ((RangeTextBox)TextBoxAttributeDic["pierLabelRange_WD"]).GetRangeFromFullAddress();
            Range storeyRange = ((RangeTextBox)TextBoxAttributeDic["matchStoreyRange_WD"]).GetRangeFromFullAddress();
            Range outputRange = ((RangeTextBox)TextBoxAttributeDic["outputRange_WD"]).GetRangeFromFullAddress();
            outputRange = outputRange.Worksheet.Cells[pierLableRange.Row,outputRange.Column];
            Range statusRange = ((RangeTextBox)TextBoxAttributeDic["statusRange_WD"]).GetRangeFromFullAddress();
            statusRange = statusRange.Worksheet.Cells[pierLableRange.Row, statusRange.Column];
            #endregion


            #region Init Arrays
            int outputLength = pierLableRange.Rows.Count;
            double[] rebarDia = new double[outputLength];
            double[] rebarSpacing = new double[outputLength];
            double[] shearDia = new double[outputLength];
            double[] shearSpacing = new double[outputLength];
            string[] status = new string[outputLength];
            #endregion

            int rowCounter = 0;
            foreach (Range cell in pierLableRange.Cells)
            {
                int rowNum = cell.Row;
                string pierLabel = cell.Value2.ToString();
                string etabsStoreyName = storeyRange.Worksheet.Cells[rowNum, storeyRange.Column].Value2.ToString();

                try
                {
                    if (!rebarDic.ContainsKey(pierLabel)) { throw new Exception($"Wall Label {pierLabel} not found in rebar table"); }
                    // update this to status instead of exception 
                    WallRebar wallRebar = rebarDic[pierLabel];

                    string[] rowValues = wallRebar.GetStoreyData(storeyTracker.GetStoreyNum(etabsStoreyName, "etabs"));
                    rebarDia[rowCounter] = double.Parse(rowValues[3]);
                    rebarSpacing[rowCounter] = double.Parse(rowValues[4]);
                    shearDia[rowCounter] = double.Parse(rowValues[5]);
                    shearSpacing[rowCounter] = double.Parse(rowValues[6]);
                    status[rowCounter] = "Completed " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
                }
                catch (Exception ex)
                {
                    status[rowCounter] = "Error: " + ex.Message;
                }
                
                rowCounter++;
            }

            #region Write To Excel
            WriteToExcelRangeAsCol(outputRange, 0, 0, false, rebarDia, rebarSpacing);
            WriteToExcelRangeAsCol(outputRange, 0, 8, false, shearDia, shearSpacing);
            WriteToExcelRangeAsCol(statusRange, 0, 0, false, status);
            #endregion

            MessageBox.Show("Completed", "Completed");
        }

        #endregion
    }

    public class WallRebar
    {
        public string Name;
        public List<string[]> tableContentsList = new List<string[]>();
        public Dictionary<int, string[]> tableContents = new Dictionary<int, string[]>();
        StoreyTracker storeyTracker;
        
        public WallRebar(string name, StoreyTracker storeyTracker) 
        {
            this.Name = name;
            this.storeyTracker = storeyTracker;
        }

        public void AddRow(Range range)
        {
            string[] rowContents = GetContentsAsStringArray(range,false);
            string startStoreyName = rowContents[1];
            int startStoreyNumber = storeyTracker.GetStoreyNum(startStoreyName, "design");
            tableContents.Add(startStoreyNumber, rowContents);
        }

        #region Sort and Find Stories
        private int[] startStoreyNumSorted = null;
        private int[] endStoreyNumSorted = null;
        public void SortStories()
        {
            startStoreyNumSorted = tableContents.Keys.OrderBy(key => key).ToArray();
            
            endStoreyNumSorted = new int[startStoreyNumSorted.Length];
            for (int i = 0; i < startStoreyNumSorted.Length; i++) 
            { 
                string endStoreyName = tableContents[startStoreyNumSorted[i]][2];
                endStoreyNumSorted[i] = storeyTracker.GetStoreyNum(endStoreyName, "design");
            }
        }


        public string[] GetStoreyData(int targetStoreyNum)
        {
            int targetIndex = -1;
            for (int i = 0; i < startStoreyNumSorted.Length; i++)
            { 
                if (targetStoreyNum >= startStoreyNumSorted[i] && targetStoreyNum <= endStoreyNumSorted[i])
                {
                    return tableContents[startStoreyNumSorted[i]];
                }
            }

            throw new Exception("Unable to find target storey");
        }
        #endregion
    }
    public class StoreyTracker
    {
        Dictionary<string, int> etabsStoreyDicNameToNum = new Dictionary<string, int>();
        Dictionary<string, int> designStoreyDicNameToNum = new Dictionary<string, int>();
        Dictionary<int, string> etabsStoreyDicNumToName = new Dictionary<int, string>();
        Dictionary<int, string> designStoreyDicNumToName = new Dictionary<int, string>();
        public StoreyTracker(Range storeyTable)
        {
            GetStoreyTable(storeyTable);
        }
        private void GetStoreyTable(Range storeyTable)
        {
            foreach (Range range in storeyTable.Rows)
            {
                int storeyNumber = Int32.Parse(range.Cells[1].Value2.ToString());
                string etabsStoreyName = range.Cells[2].Value2.ToString();
                string designStoreyName = range.Cells[3].Value2.ToString();

                etabsStoreyDicNameToNum.Add(etabsStoreyName, storeyNumber);
                designStoreyDicNameToNum.Add(designStoreyName, storeyNumber);
                etabsStoreyDicNumToName.Add(storeyNumber, etabsStoreyName);
                designStoreyDicNumToName.Add(storeyNumber, designStoreyName);
            }
        }

        public int GetStoreyNum(string name, string type)
        {
            if (type == "etabs")
            {
                if (!etabsStoreyDicNameToNum.ContainsKey(name)) { throw new Exception("ETABS Storey Name not found in reference table."); }
                return etabsStoreyDicNameToNum[name];
            }
            else if (type == "design")
            {
                if (!designStoreyDicNameToNum.ContainsKey(name)) { throw new Exception("Design Storey Name not found in reference table."); }
                return designStoreyDicNameToNum[name];
            }
            else
            {
                throw new Exception($"Invalid type {type} for GetStoreyNum");
            }
        }

        public string GetStoreyName(int num, string type)
        {
            if (type == "etabs")
            {
                if (!etabsStoreyDicNumToName.ContainsKey(num)) { throw new Exception("ETABS Storey Number not found in reference table."); }
                return etabsStoreyDicNumToName[num];
            }
            else if (type == "design")
            {
                if (!designStoreyDicNumToName.ContainsKey(num)) { throw new Exception("Design Storey Number not found in reference table."); }
                return designStoreyDicNumToName[num];
            }
            else
            {
                throw new Exception($"Invalid type {type} for GetStoreyNum");
            }
        }
    }
}
