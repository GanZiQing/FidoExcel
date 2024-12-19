using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class BeamDesign : UserControl
    {
        #region Init
        public BeamDesign()
        {
            InitializeComponent();
            CreateAttributes();
        }
        Dictionary<string, AttributeTextBox> RangeAttributeDic = new Dictionary<string, AttributeTextBox>();
        private void CreateAttributes()
        {
            #region Interpolation
            var thisAtt = new RangeTextBox("beamTable_beam", dispBeamTable, setBeamTable, "range");
            RangeAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new RangeTextBox("outputCol_beam", dispOutputColumn, setOutputColumn, "cell");
            RangeAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new RangeTextBox("shearTable_beam", dispShearTable, setShearTable, "range");
            RangeAttributeDic.Add(thisAtt.attName, thisAtt);
            #endregion
        }
        #endregion

        #region Decompose Table
        private void decomposeTable_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {
                try
                {
                    #region Get Ranges
                    Range beamTable = ((RangeTextBox)RangeAttributeDic["beamTable_beam"]).GetRangeFromFullAddress();
                    Range outputRange = ((RangeTextBox)RangeAttributeDic["outputCol_beam"]).GetRangeFromFullAddress();
                    Range shearTable = ((RangeTextBox)RangeAttributeDic["shearTable_beam"]).GetRangeFromFullAddress();
                    Worksheet sheet = beamTable.Worksheet;
                    #endregion

                    #region Loop through Data
                    (int startRow, int endRow, int startCol, int endCol) = CommonUtilities.GetRangeDetails(beamTable);

                    (int writeStartRow, _, int writeStartCol, _) = CommonUtilities.GetRangeDetails(outputRange);
                    int rowCounter = 0;
                    for (int rowNum = startRow; rowNum <= endRow; rowNum++)
                    {
                        try
                        {
                            #region Update Progress
                            if (worker.CancellationPending) { return; }
                            progressTracker.UpdateStatus($"Checking items for row {rowNum}. {rowCounter} / {endRow - startRow + 1} rows completed.");
                            worker.ReportProgress(CommonUtilities.ConvertToProgress(rowCounter, endRow - startRow + 1)); 
                            rowCounter++;
                            #endregion

                            #region Checks
                            if (sheet.Cells[rowNum, startCol].Value2 == null) { continue; }
                            string checkContent = (string)sheet.Cells[rowNum, startCol].Value2.ToString();
                            checkContent = checkContent.Substring(0, 1);
                            bool isValidRow = Int32.TryParse(checkContent, out _);
                            if (!isValidRow) { continue; }
                            #endregion

                            #region Top Reinforcement
                            Range T1 = sheet.Cells[rowNum, startCol];
                            int a1 = T1.Row;
                            int a2 = T1.Column;
                            (string[] Tlrebars, double T1As) = SplitRebar(T1);
                            Range T3 = sheet.Cells[rowNum, startCol + 2];
                            (string[] T3rebars, double T3As) = SplitRebar(T3);

                            string[] TRebar;
                            if (T1As < T3As)
                            {
                                TRebar = Tlrebars;
                            }
                            else
                            {
                                TRebar = T3rebars;
                            }
                            #endregion

                            #region Bottom Reinforcement
                            Range B2 = sheet.Cells[rowNum, startCol + 4];
                            (string[] B2rebar, double B2As) = SplitRebar(B2);
                            #endregion

                            #region Shear Reinforcement
                            Range S1 = sheet.Cells[rowNum, startCol + 6];
                            int[] S1rebar = GetLink(S1, shearTable);
                            Range S3 = sheet.Cells[rowNum, startCol + 8];
                            int[] S3rebar = GetLink(S1, shearTable);

                            int[] SRebar = new int[3];
                            if (S1rebar[3] > S3rebar[3])
                            {
                                Array.Copy(S1rebar, SRebar, S1rebar.Length - 1);
                            }
                            else
                            {
                                Array.Copy(S3rebar, SRebar, S1rebar.Length - 1);
                            }
                            #endregion

                            #region Write Result
                            string[] writeArray = CommonUtilities.ConcatArrays(new List<Array> { TRebar, B2rebar, SRebar });
                            CommonUtilities.WriteToExcelRangeAsRow(sheet.Cells[rowNum, writeStartCol], 0, 0, false, writeArray);
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            sheet.Cells[rowNum, writeStartCol].Value2 = "Error:" + ex.Message;
                        }
                    }
                    #endregion
                    MessageBox.Show("Completed", "Completed");
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
            });
        }
        private (string[] rebars, double As) SplitRebar(Range range)
        {
            string T1Contents = range.Value2.ToString();
            T1Contents = T1Contents.Trim();
            string[] rebarLayers = T1Contents.Split('+');
            for (int i = 0; i< rebarLayers.Length; i++) { rebarLayers[i] = rebarLayers[i].Trim(); }

            string[] rebarColToPrint = new string[4];
            if (rebarLayers.Length == 1 || rebarLayers.Length == 2)
            { 
            string[] layer1 = rebarLayers[0].Split('H');
            rebarColToPrint[0] = layer1[0].Trim();
            rebarColToPrint[1] = layer1[1].Trim();
            }

            if (rebarLayers.Length == 2)
            {
                string[] layer2 = rebarLayers[1].Split('H');
                rebarColToPrint[2] = layer2[0].Trim();
                rebarColToPrint[3] = layer2[1].Trim();
            }
            else if (rebarLayers.Length > 2)
            {
                SplitRebarByDiameter(rebarLayers, ref rebarColToPrint, range);
                //throw new ArgumentException($"More than 2 layers encountered in cell {range.Row}, {range.Column}");
            }

            #region Calculate As
            double As = CalculateAs(rebarColToPrint);
            #endregion
            return (rebarColToPrint, As);
        }

        private void SplitRebarByDiameter(string[] rebarLayers, ref string[] rebarColToPrint, Range range)
        {
            #region Get Diameters
            Dictionary<int,int> rebarsDict = new Dictionary<int,int>();            
            foreach (string layer in rebarLayers)
            {
                string[] layerSplit = layer.Split('H');
                int num = Int32.Parse(layerSplit[0].Trim());
                int dia = Int32.Parse(layerSplit[1].Trim());
                if (!rebarsDict.ContainsKey(dia))
                {
                    rebarsDict[dia] = num;
                }
                else
                {
                    rebarsDict[dia] += num;
                }
            }
            #endregion

            #region Sort add to print
            if (rebarsDict.Count > 2) { throw new ArgumentException($"More than 2 types of rebar encountered in cell {range.Row}, {range.Column}"); }

            int currentIndex = 0;

            var descendingDiameters = rebarsDict.Keys.OrderByDescending(key => key);
            foreach (int diameterKey in descendingDiameters)
            {
                rebarColToPrint[currentIndex] = rebarsDict[diameterKey].ToString();
                rebarColToPrint[currentIndex + 1] = diameterKey.ToString();
                currentIndex += 2;
            }
            #endregion
        }

        private double CalculateAs(string[] rebar)
        { 
            double As = 0;
            for (int i = 0; i < rebar.Length/2; i++)
            {
                if (rebar[i*2] == null) { break; }
                double n = double.Parse(rebar[i*2]);
                double d = double.Parse(rebar[i*2+1]);
                As += n * (Math.PI * Math.Pow(d, 2)) / 4;
            }
            return As;
        }

        private int[] GetLink(Range range, Range shearTable)
        {
            string contents = range.Value2;
            contents = contents.Trim();
            int numBar = 1;
            if(contents.Length > 4)
            {
                numBar = Int32.Parse(contents.Substring(0,1));
                contents = contents.Substring(1,contents.Length-1);
            }

            #region Find Max As
            (int dia, int spacing, double shearAs) = FindLinkFromTable(contents, shearTable);

            int[] linkDetails = new int[4];
            linkDetails[0] = numBar;
            linkDetails[1] = dia;
            linkDetails[2] = spacing;
            linkDetails[3] = (int)(shearAs*(double)numBar);
            #endregion

            return linkDetails;
        }

        private (int dia, int spacing, double shearAs) FindLinkFromTable(string content, Range table)
        {
            (int startRow, int endRow, int startCol, int endCol) = CommonUtilities.GetRangeDetails(table);
            Worksheet sheet = table.Worksheet;
            for (int rowNum = 1; rowNum <= table.Rows.Count; rowNum++) 
            {
                Range cell = table.Cells[rowNum, 1];
                if (cell.Value2 == null) { continue; }

                if (cell.Value2.ToString() == content)
                {
                    int dia = (int)table.Cells[rowNum, 3].Value2;
                    int spacing = (int)table.Cells[rowNum, 4].Value2;
                    double shearAs = table.Cells[rowNum, 5].Value2;
                    return (dia, spacing, shearAs);
                }
            }
            throw new ArgumentException($"Unable to find {content} in shear link table");
        }

        #endregion

        private void setBeamTable_Click(object sender, EventArgs e)
        {

        }
    }

}
