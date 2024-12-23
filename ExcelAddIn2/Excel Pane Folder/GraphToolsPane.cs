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
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.Excel.Application;
//using static ExcelAddIn2.Piling.Utilities;
using Microsoft.VisualBasic;
using SeriesCollection = Microsoft.Office.Interop.Excel.SeriesCollection;
using static ExcelAddIn2.CommonUtilities;
using System.Windows.Forms.Integration;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class GraphToolsPane : UserControl
    {
        Application thisApp = Globals.ThisAddIn.Application;
        Dictionary<string, AttributeTextBox> RangeAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();
        public GraphToolsPane()
        {
            InitializeComponent();
            CreateAttributes();
            //ImportTaskPane();
        }

        private void CreateAttributes()
        {
            #region Chart Tools
            RangeTextBox xRange_chart = new RangeTextBox("xRange_chart", dispXRange, setXRange, "range");
            RangeAttributeDic.Add("xRange_chart", xRange_chart);

            RangeTextBox yRange_chart = new RangeTextBox("yRange_chart", dispYRange, setYRange, "cell");
            RangeAttributeDic.Add("yRange_chart", yRange_chart);


            RangeTextBox nameRange_chart = new RangeTextBox("nameRange_chart", dispNameRange, setNameRange, "range");
            RangeAttributeDic.Add("nameRange_chart", nameRange_chart);

            CustomAttribute thisCustomAtt = new CheckBoxAttribute("terminateAtNull_chart", terminateAtNullCheck, true);
            OtherAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisCustomAtt = new CheckBoxAttribute("plotPoints_chart", pointCheck, false);
            OtherAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisCustomAtt = new CheckBoxAttribute("plotLines_chart", lineCheck, false);
            OtherAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisCustomAtt = new CheckBoxAttribute("clearChart_chart", clearChartCheck, true);
            OtherAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);
            #endregion


            #region Interpolation
            thisCustomAtt = new MultipleRangeAttribute("interpolationRange_chart", setRange1, true);
            OtherAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            AttributeTextBox thisAtt = new RangeTextBox("outputRange_chart", dispOutputRange, setOutputRange, "cell");
            RangeAttributeDic.Add(thisAtt.attName, thisAtt);

            thisAtt = new RangeTextBox("dataSeriesRange_chart", dispDataSeries, setDataSeries, "range");
            RangeAttributeDic.Add(thisAtt.attName, thisAtt);

            thisCustomAtt = new CheckBoxAttribute("terminateAtNull2_chart", terminateAtNullCheck2, true);
            OtherAttributeDic.Add(thisCustomAtt.attName, thisCustomAtt);

            thisAtt = new AttributeTextBox("ampFactor_chart", dispAmpFactor, true);
            thisAtt.type = "double";
            thisAtt.SetDefaultValue("0");
            RangeAttributeDic.Add(thisAtt.attName, thisAtt);
            #endregion

            #region Upgrade Interpolation

            #endregion
        }

        private void ImportTaskPane()
        {
            PilingPane RefPane = new PilingPane();
            tabControl.TabPages.Add(RefPane.GetPageTaskPane(0));
            tabControl.TabPages.Add(RefPane.GetPageTaskPane(0));
        }

        #region Chart
        
        private Chart GetChart()
        {
            Chart chart = null;
            chart = thisApp.ActiveChart;
            if (chart == null)
            {
                throw new Exception($"No chart object found");
            }
            return chart;
        }

        private void addSeries_Click(object sender, EventArgs e)
        {
            ProgressHelper.RunWithProgress((worker, progressTracker) =>
            {

                #region Get Excel Detail
                Range xRange;  Range yRange; Range nameRange;
                try
                {
                    progressTracker.UpdateStatus("Getting Excel Data");
                    xRange = ((RangeTextBox)RangeAttributeDic["xRange_chart"]).GetRangeFromFullAddress();
                    TerminateRangeAtNullFirstCell(ref xRange);
                    {
                        Range yCell = ((RangeTextBox)RangeAttributeDic["yRange_chart"]).GetRangeFromFullAddress();
                        yRange = GetColRangeFromRanges(xRange, yCell);
                        Range nameCell = ((RangeTextBox)RangeAttributeDic["nameRange_chart"]).GetRangeFromFullAddress();
                        nameRange = GetColRangeFromRanges(xRange, nameCell);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error getting excel detail\n{ex.Message}");
                }

                (int startRow, int endRow, int startCol, int endCol) = GetRangeDetails(xRange);
                #endregion

                #region Get Chart
                SeriesCollection seriesCollection = null;
                Chart chart = null;
                try
                {
                    chart = GetChart();
                    seriesCollection = chart.SeriesCollection();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error getting chart");
                    return;
                }
                #endregion

                #region Clear chart
                if (clearChartCheck.Checked)
                {
                    while (seriesCollection.Count > 0)
                    {
                        chart.SeriesCollection(1).Delete();
                    }
                    seriesCollection = chart.SeriesCollection();
                }
                #endregion

                #region Plotting
                progressTracker.UpdateStatus("Plotting...");
                int currentRow = 1;
                Series newSeries;
                while (currentRow <= xRange.Rows.Count)
                {
                    // Handle last row
                    if (currentRow == xRange.Rows.Count) 
                    {
                        newSeries = seriesCollection.NewSeries();
                        newSeries.XValues = xRange[currentRow];
                        newSeries.Values = yRange[currentRow];
                        newSeries.Name = $"= {nameRange[currentRow].Worksheet.Name}!{nameRange[currentRow].Address[false, false]}";
                        break;
                    }

                    // Find the end of the current series
                    endRow = currentRow;
                    while (endRow <= xRange.Rows.Count - 1)
                    {
                        if (nameRange[currentRow].Value2 != nameRange[endRow + 1].Value2)
                        {
                            break;
                        }
                        endRow += 1;
                    }

                    // Plot series
                    newSeries = seriesCollection.NewSeries();
                    newSeries.XValues = xRange.Worksheet.Range[xRange[currentRow], xRange[endRow]];
                    newSeries.Values = yRange.Worksheet.Range[yRange[currentRow], yRange[endRow]];
                    string testname = $"= '{nameRange[currentRow].Worksheet.Name}'!{nameRange[currentRow].Address[false, false]}";
                    newSeries.Name = $"= '{nameRange[currentRow].Worksheet.Name}'!{nameRange[currentRow].Address[false, false]}";

                    // Increment
                    worker.ReportProgress(ConvertToProgress(currentRow, xRange.Rows.Count));
                    currentRow = endRow + 1;
                }
                #endregion

                #region Format Plot
                progressTracker.UpdateStatus("Formatting plot");
                if (pointCheck.Checked && lineCheck.Checked)
                {
                    chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines;
                }
                else if (lineCheck.Checked)
                {
                    chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLinesNoMarkers;
                }
                else if (pointCheck.Checked)
                {
                    chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                }

                chart.Refresh();
                progressTracker.UpdateStatus("Completed");
                #endregion
                MessageBox.Show("Completed", "Completed");


            });
        }

        private void clearChart_Click(object sender, EventArgs e)
        {
            #region Get Chart
            SeriesCollection seriesCollection = null;
            Chart chart = null;
            try
            {
                chart = GetChart();
                seriesCollection = chart.SeriesCollection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error getting chart");
                return;
            }
            #endregion

            while (seriesCollection.Count > 0)
            {
                chart.SeriesCollection(1).Delete();
            }
        }
        #endregion

        #region Interpolation
        private void runInterpolation_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Ranges
                (_, Range[] selectedRanges) = ((MultipleRangeAttribute)OtherAttributeDic["interpolationRange_chart"]).GetRanges();
                Range outputCell = ((RangeTextBox)RangeAttributeDic["outputRange_chart"]).GetRangeFromFullAddress();
                double amplificationFactor = RangeAttributeDic["ampFactor_chart"].GetDoubleFromTextBox();

                if (terminateAtNullCheck2.Checked)
                {
                    for (int i = 0; i < selectedRanges.Length; i++)
                    {
                        TerminateRangeAtNullFirstCell(ref selectedRanges[i]);
                    }
                }
                // Convert into double lists
                List<double[]> rangesX = new List<double[]>();
                List<double[]> rangesY = new List<double[]>();
                foreach (Range range in selectedRanges)
                {
                    if (range.Columns.Count != 2)
                    {
                        throw new ArgumentException($"Range provided does not have 2 columns. \nRange provided: {range.Address[false, false]}");
                    }
                    rangesX.Add(GetContentsAsDoubleArray(range.Columns[1].Cells));
                    rangesY.Add(GetContentsAsDoubleArray(range.Columns[2].Cells));
                }
                #endregion

                #region Check is ascending/descending
                bool isAscending = CheckAllAscendingOrDescending(rangesX);
                #endregion

                #region Create new X values
                double[] finalX = CombineAllRanges(rangesX, isAscending);
                #endregion

                #region Create new Y values
                List<double[]> finalYs = new List<double[]>();
                for (int i = 0; i < rangesX.Count; i++)
                {
                    double[] rangeX = rangesX[i];
                    double[] rangeY = rangesY[i];
                    double[] finalY = InterpolateValues(finalX, rangeX, rangeY, isAscending);
                    finalYs.Add(finalY);
                }
                #endregion

                #region Create SuperImposed Values 
                double[] finalSuperY = new double[finalX.Length];
                
                for (int i = 0; i < finalX.Length; i++)
                {
                    foreach(double[] finalY in finalYs)
                    {
                        finalSuperY[i] += finalY[i];
                    }
                }
                finalYs.Add(finalSuperY);

                if (amplificationFactor != 0)
                {
                    double[] finalSuperY2 = new double[finalX.Length];
                    for (int i = 0; i < finalX.Length; i++)
                    {
                        finalSuperY2[i] = finalSuperY[i] * amplificationFactor;
                    }
                    finalYs.Add(finalSuperY2);
                }
                
                #endregion

                #region Get Series Names
                List<string[]> seriesNames = GetSeriesName(finalX, finalYs);
                #endregion

                #region Write Series to Excel
                outputCell.Worksheet.Activate();
                outputCell.Select();
                for (int i = 0; i < finalYs.Count; i++)
                {
                    WriteToExcelSelectionAsRow(0, i * 4, false, seriesNames[i], finalX, finalYs[i]);
                }
                #endregion

                #region Write Stacked Series to Excel
                int rowOffset = 0;
                int colOffset = finalYs.Count * 4;

                for (int i = 0; i < finalYs.Count; i++)
                {
                    WriteToExcelSelectionAsRow(rowOffset, colOffset, false, seriesNames[i], finalX, finalYs[i]);
                    rowOffset += finalYs[i].Length;
                }
                #endregion

                MessageBox.Show("Completed", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void runInterpolation_ClickArchive(object sender, EventArgs e)
        {
            //#region Get Excel Ranges
            //Range range1 = ((RangeTextBox)RangeAttributeDic["range1_chart"]).GetRangeFromFullAddress();
            //Range range2 = ((RangeTextBox)RangeAttributeDic["range2_chart"]).GetRangeFromFullAddress();
            //Range outputCell = ((RangeTextBox)RangeAttributeDic["outputRange_chart"]).GetRangeFromFullAddress();
            //if (terminateAtNullCheck2.Checked)
            //{
            //    TerminateColRangeAtNull(ref range1);
            //    TerminateColRangeAtNull(ref range2);
            //}
            //double[] range1X = GetContentsAsDoubleArray(range1.Columns[1].Cells);
            //double[] range1Y = GetContentsAsDoubleArray(range1.Columns[2].Cells);
            //double[] range2X = GetContentsAsDoubleArray(range2.Columns[1].Cells);
            //double[] range2Y = GetContentsAsDoubleArray(range2.Columns[2].Cells);
            //#endregion

            //#region Check is ascending/descending
            //bool isAscending;
            //try
            //{
            //    bool isAscending1X = CheckArrayAscendingOrDescending(range1X);
            //    bool isAscending1Y = CheckArrayAscendingOrDescending(range2X);
            //    if (isAscending1X != isAscending1Y)
            //    {
            //        throw new Exception($"Range 1 and 2 are sorted in different orders");
            //    }
            //    isAscending = isAscending1X;
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Error");
            //    return;
            //}
            
            
            //#endregion

            //#region Create new X values
            //double[] finalX = CombineTwoRangesOG(range1X, range2X, isAscending);
            //#endregion

            //#region Create new Y values
            //double[] final1Y = InterpolateValues(finalX, range1X, range1Y, isAscending);
            //double[] final2Y = InterpolateValues(finalX, range2X, range2Y, isAscending);
            //#endregion

            //#region Find superimposed values
            //double[] finalSuperY = new double[finalX.Length];
            //for (int i = 0; i < finalX.Length; i++)
            //{
            //    finalSuperY[i] = final1Y[i] + final2Y[i];
            //}
            //#endregion

            //#region Get Series Names
            //string[] inputSeriesNames = GetContentsAsStringArray(((RangeTextBox)RangeAttributeDic["dataSeriesRange_chart"]).GetRangeFromFullAddress(), true);


            //List<string[]> ListForPlot = new List<string[]>();
            //for (int intputSeriesIndex = 0; intputSeriesIndex < inputSeriesNames.Length; intputSeriesIndex++)
            //{
            //    string[] seriesName = new string[finalX.Length];
            //    for (int seriesNameIndex = 0; seriesNameIndex < seriesName.Length; seriesNameIndex++)
            //    {
            //        seriesName[seriesNameIndex] = inputSeriesNames[intputSeriesIndex];
            //    }
            //    ListForPlot.Add(seriesName);
            //}
            //#endregion

            //#region Write to Excel
            //outputCell.Select();
            //WriteToExcel(0, 0, false, ListForPlot[0], finalX, final1Y);
            //WriteToExcel(0, 4, false, ListForPlot[1], finalX, final2Y);
            //WriteToExcel(0, 8, false, ListForPlot[2], finalX, finalSuperY);
            //#endregion

            //#region Write ULS
            //double[] ulsY = new double[finalSuperY.Length];
            //for (int i =0; i < ulsY.Length; i++)
            //{
            //    ulsY[i] = finalSuperY[i] * 1.4;
            //}
            //WriteToExcel(0, 12, false, ListForPlot[3], finalX, ulsY);
            //#endregion

            //#region Write Collated Array
            //int rowNum = 0;
            //WriteToExcel(rowNum, 16, false, ListForPlot[0], finalX, final1Y);
            //rowNum += finalX.Length;
            //WriteToExcel(rowNum, 16, false, ListForPlot[1], finalX, final2Y);
            //rowNum += finalX.Length;
            //WriteToExcel(rowNum, 16, false, ListForPlot[2], finalX, finalSuperY);
            //rowNum += finalX.Length;
            //WriteToExcel(rowNum, 16, false, ListForPlot[3], finalX, ulsY);
            //#endregion
            //MessageBox.Show("Completed", "Completed");
        }
        private Dictionary<double, List<double>> RangeToDictionary(double[] rangeKey, double[] rangeValue)
        {
            // Probably should rewrite to take any type as input
            if (rangeKey.Length != rangeValue.Length)
            {
                throw new Exception($"Unable to convert ranges as their lengths are unequal.");
            }
            Dictionary<double, List<double>> output = new Dictionary<double, List<double>>();
            for (int i = 0; i < rangeKey.Length; i++)
            {
                if (!output.ContainsKey(rangeKey[i]))
                {
                    List<double> value = new List<double>();
                    output.Add(rangeKey[i], value);   
                }
                output[rangeKey[i]].Add(rangeValue[i]);
            }

            return output;
        }

        private Dictionary<double, double> RangeToDictionaryOG(double[] rangeKey, double[] rangeValue)
        {
            // Probably should rewrite to take any type as input
            if (rangeKey.Length != rangeValue.Length)
            {
                throw new Exception($"Unable to convert ranges as their lengths are unequal.");
            }
            Dictionary<double, double> output = new Dictionary<double, double>();
            for (int i = 0; i< rangeKey.Length; i++)
            {
                output.Add(rangeKey[i], rangeValue[i]);
            }

            return output;
        }

        private double[] CombineTwoRanges(double[] range1X, double[] range2X, bool isAscending)
        {
            int r1Counter = 0;
            int r2Counter = 0;
            List<double> finalXList = new List<double>();

            while (r1Counter < range1X.Length || r2Counter < range2X.Length)
            {
                double value1;
                double value2;
                #region Check if any index has been exceeded
                if (r1Counter >= range1X.Length)
                {
                    value2 = range2X[r2Counter];
                    finalXList.Add(value2);
                    r2Counter++;
                    continue;
                }
                else if (r2Counter >= range2X.Length)
                {
                    value1 = range1X[r1Counter];
                    finalXList.Add(value1);
                    r1Counter++;
                    continue;
                }
                #endregion

                value1 = range1X[r1Counter];
                value2 = range2X[r2Counter];


                #region Check if any value is unplotable
                if (double.IsNaN(value1))
                {
                    r1Counter++;
                    continue;
                }
                if (double.IsNaN(value2))
                {
                    r2Counter++;
                    continue;
                }
                #endregion

                if (value1 == value2)
                {
                    finalXList.Add(value1);
                    r1Counter++;
                    r2Counter++;
                }
                else if (value1 < value2)
                {
                    if (isAscending)
                    {
                        finalXList.Add(value1);
                        r1Counter++;
                    }
                    else
                    {
                        finalXList.Add(value2);
                        r2Counter++;
                    }
                }
                else if (value1 > value2)
                {
                    if (isAscending)
                    {
                        finalXList.Add(value2);
                        r2Counter++;
                    }
                    else
                    {
                        finalXList.Add(value1);
                        r1Counter++;
                    }
                }
                else
                {
                    throw new Exception($"Unable to compare values: {value1}, {value2}");
                }
            }
            return finalXList.ToArray();
        }

        private double[] CombineAllRanges(List<double[]> rangesX, bool isAscending)
        {
            double[] finalRangeX = rangesX[0];
            for (int i = 1; i < rangesX.Count; i++)
            {
                double[] addRange = rangesX[i];
                finalRangeX = CombineTwoRanges(finalRangeX, addRange, isAscending);
            }
            return finalRangeX;
        }

        private double[] InterpolateValues(double[] finalX, double[] rangeX, double[] rangeY, bool isAscending)
        {
            Dictionary<double, List<double>> range1Dic = RangeToDictionary(rangeX, rangeY);
            Dictionary<double, int> range1MultiValueIndex = new Dictionary<double, int>();// this keeps track of which values have already been extracted
            double[] finalY = new double[finalX.Length];

            double minX = rangeX.Min();
            double maxX = rangeX.Max();

            for (int iFinal = 0; iFinal < finalX.Length; iFinal++)
            {
                double x = finalX[iFinal];

                #region Check if X within bounds
                if (x < minX || x > maxX)
                {
                    finalY[iFinal] = 0;
                    continue;
                }
                #endregion

                #region If X already exist
                if (range1Dic.Keys.Contains(x))
                {
                    List<double> values = range1Dic[x];
                    if (values.Count == 1)
                    {
                        finalY[iFinal] = values[0];
                    }
                    else
                    {
                        if (!range1MultiValueIndex.ContainsKey(x))
                        {
                            range1MultiValueIndex[x] = 0;
                        }

                        finalY[iFinal] = values[range1MultiValueIndex[x]];
                        range1MultiValueIndex[x]++;
                    }
                    continue;
                }
                #endregion

                #region Find closest index
                int closestIndex = -1;
                for (int rangeXIndex = 0; rangeXIndex < rangeX.Length; rangeXIndex++)
                {
                    if (isAscending)
                    {
                        if (rangeX[rangeXIndex] > x)
                        {
                            closestIndex = rangeXIndex - 1;
                            break;
                        }
                    }
                    else
                    {
                        if (rangeX[rangeXIndex] < x)
                        {
                            closestIndex = rangeXIndex - 1;
                            break;
                        }
                    }
                }                
                #endregion

                #region Find gradients and c
                double gradient = (rangeY[closestIndex] - rangeY[closestIndex + 1]) / (rangeX[closestIndex] - rangeX[closestIndex + 1]);
                double c = rangeY[closestIndex] - gradient * rangeX[closestIndex];
                #endregion

                #region Find Y value7
                finalY[iFinal] = gradient * x + c;
                #endregion
            }
            return finalY;
        }

        private double[] InterpolateValuesOG(double[] finalX, double[] rangeX, double[] rangeY, bool isAscending)
        {
            Dictionary<double, double> range1Dic = RangeToDictionaryOG(rangeX, rangeY);
            double[] finalY = new double[finalX.Length];

            double minX = rangeX.Min();
            double maxX = rangeX.Max();

            for (int iFinal = 0; iFinal < finalX.Length; iFinal++)
            {
                double x = finalX[iFinal];

                #region Check if X within bounds
                if (x < minX || x > maxX)
                {
                    finalY[iFinal] = 0;
                    continue;
                }
                #endregion

                #region If X already exist
                if (range1Dic.Keys.Contains(x))
                {
                    finalY[iFinal] = range1Dic[x];
                    continue;
                }
                #endregion

                #region Find closest index
                int closestIndex = -1;
                for (int rangeXIndex = 0; rangeXIndex < rangeX.Length; rangeXIndex++)
                {
                    if (isAscending)
                    {
                        if (rangeX[rangeXIndex] > x)
                        {
                            closestIndex = rangeXIndex - 1;
                            break;
                        }
                    }
                    else
                    {
                        if (rangeX[rangeXIndex] < x)
                        {
                            closestIndex = rangeXIndex - 1;
                            break;
                        }
                    }
                }
                #endregion

                #region Find gradients and c
                double gradient = (rangeY[closestIndex] - rangeY[closestIndex + 1]) / (rangeX[closestIndex] - rangeX[closestIndex + 1]);
                double c = rangeY[closestIndex] - gradient * rangeX[closestIndex];
                #endregion

                #region Find Y value
                finalY[iFinal] = gradient * x + c;
                #endregion
            }
            return finalY;
        }

        private bool CheckAllAscendingOrDescending(List<double[]> rangesX)
        {
            bool isAscending = CheckArrayAscendingOrDescending(rangesX[0]);
            
            foreach (double[] rangeX in rangesX)
            {
                bool isThisAscending = CheckArrayAscendingOrDescending(rangeX);
                if (isThisAscending != isAscending)
                {
                    throw new Exception($"Ranges are sorted in different orders");
                }
            }
            return isAscending;
        }
        private bool CheckArrayAscendingOrDescending(double[] array)
        {
            bool isAscending = true;
            bool isDescending = true;

            for (int i = 0; i < array.Length - 1; i++)
            {
                if (array[i] > array[i + 1] && isAscending)
                {
                    isAscending = false;
                }
                if (array[i] < array[i + 1] && isDescending)
                {
                    isDescending = false;
                }
            }

            if (isAscending)
            {
                return true; // Array is ascending
            }

            if (isDescending)
            {
                return false; // Array is descending
            }

            throw new Exception("Array provided is neither ascending nor descending");
        }
        
        private List<string[]> GetSeriesName(double[] finalX, List<double[]> finalYs)
        {
            // Converts input array of names to big arrays for plotting
            string[] inputSeriesNames = GetContentsAsStringArray(((RangeTextBox)RangeAttributeDic["dataSeriesRange_chart"]).GetRangeFromFullAddress(), true);

            #region Check Lengths 
            if (inputSeriesNames.Length < finalYs.Count)
            {
                MessageBox.Show($"Warning:\n" +
                    $"Number of series names provided is insufficient.\n" +
                    $"Expected number of series names is {finalYs.Count}, number of series names provided is {inputSeriesNames.Length}\n" +
                    $"Default series names used.");
            }
            else if (inputSeriesNames.Length > finalX.Length)
            {
                MessageBox.Show($"Warning:\n" +
                    $"Number of series names provided exceeds final data range.\n" +
                    $"Expected number of series names is {finalYs.Count}, number of series names provided is {inputSeriesNames.Length}\n");
            }
            #endregion

            List<string[]> seriesNamesArrays = new List<string[]>();
            int addSeriesCounter = 1;
            for (int i = 0; i < finalYs.Count; i++)
            {
                string[] seriesNameArray;
                if (i < inputSeriesNames.Length)
                {
                    seriesNameArray = Enumerable.Repeat(inputSeriesNames[i], finalX.Length).ToArray();
                }
                else
                {
                    seriesNameArray = Enumerable.Repeat($"Series {addSeriesCounter}", finalX.Length).ToArray();
                    addSeriesCounter++;
                }
                seriesNamesArrays.Add(seriesNameArray);
            }
            return seriesNamesArrays;
        }
        #endregion
    }
}
