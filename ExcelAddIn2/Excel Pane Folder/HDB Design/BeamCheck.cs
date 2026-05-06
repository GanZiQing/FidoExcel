using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;

namespace ExcelAddIn2.Excel_Pane_Folder.HDB_Design
{
    public partial class BeamCheck : UserControl
    {
        Dictionary<string, object> attributeDic = new Dictionary<string, object>();
        public BeamCheck()
        {
            InitializeComponent(); 
            CreateAttributes();
            AddToolTips();
        }
        private void CreateAttributes()
        {
            #region Convert Rebar
            AttributeTextBox attTB;
            attTB = new RangeTextBox("beamAsIn_BC", dispAsInRange, setAsInRange, "range", false);
            attributeDic.Add(attTB.attName, attTB);
            attTB = new RangeTextBox("beamAsOut_BC", dispAsOutRange, setAsOutRange, "cell", false);
            attributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("beamAsvIn_BC", dispAsvInRange, setAsvInRange, "range", false);
            attributeDic.Add(attTB.attName, attTB);
            attTB = new RangeTextBox("beamAsvOut_BC", dispAsvOutRange, setAsvOutRange, "cell", false);
            attributeDic.Add(attTB.attName, attTB);
            #endregion
        }
        private void AddToolTips()
        {
            //toolTip1.SetToolTip(overwriteRebarCheck,
            //    "If unchecked, initial check will be done based on current values in the output range\n" +
            //    "If checked, initial check will be done based on values matched from Rebar Table");
        }
        #region Convert Main Rebar
        private object ConvertRebarStringToAs(string rebarString)
        {
            try
            {
                string[] rebarLayers = rebarString.Split('+');
                double asProvTotal = 0;
                foreach (string rebarLayer in rebarLayers)
                {
                    string[] parts = rebarLayer.Trim().Split('H');
                    if (parts.Length != 2) { throw new Exception($"Unable to split rebar layer ${rebarString}, format expected:XXHXX"); }
                    double num = double.Parse(parts[0]);
                    double dia = double.Parse(parts[1]);
                    double asProv = num * Math.PI * Math.Pow(dia, 2) / 4;
                    asProvTotal += asProv;
                }
                int asProvTotalInt = (int)Math.Floor(asProvTotal);
                return asProvTotalInt;
            }
            catch (Exception ex)
            {
                return $"Error converting rebar\n{ex.Message}";
            }            
        }
        private void convertToAs_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Data
                Range sourceRange = ((RangeTextBox)attributeDic["beamAsIn_BC"]).GetRangeForCurrentSheet();
                object[,] rebarProv = GetContentsAsObject2DArray(sourceRange);
                #endregion

                #region Conversion
                object[,] asProv = new object[rebarProv.GetLength(0), rebarProv.GetLength(1)];
                for (int rowNum = 0; rowNum < rebarProv.GetLength(0); rowNum++)
                {
                    for (int colNum = 0; colNum < rebarProv.GetLength(1); colNum++)
                    {
                        if (rebarProv[rowNum, colNum] != null)
                        {
                            string rebarStr = rebarProv[rowNum, colNum].ToString();
                            asProv[rowNum, colNum] = ConvertRebarStringToAs(rebarStr);
                        }
                    }
                }
                #endregion

                #region Write to Excel
                Range writeRange = ((RangeTextBox)attributeDic["beamAsOut_BC"]).GetRangeForCurrentSheet();
                WriteObjectToExcelRange(writeRange, 0, 0, true, asProv);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #endregion

        #region Convert Shear Rebar
        private object ConvertRebarStringToAsv(string rebarString)
        {
            try
            {
                string[] rebarLayers = rebarString.Split('+');
                double asvProvTotal = 0;
                foreach (string rebarLayer in rebarLayers)
                {
                    string[] parts = rebarLayer.Trim().Split('H');
                    if (parts.Length != 2) { throw new Exception($"Unable to split rebar layer ${rebarString}, format expected:XXHXX-XXX"); }
                    double num = double.Parse(parts[0]);

                    string[] parts2 = parts[1].Split('-');
                    if (parts2.Length != 2) { throw new Exception($"Unable to split rebar layer ${rebarString}, format expected:XXHXX-XXX"); }
                    double dia = double.Parse(parts2[0]);
                    double spacing = double.Parse(parts2[1]);
                    double asvProv = num * Math.PI * Math.Pow(dia, 2) / 4 * (1000/spacing);
                    asvProvTotal += asvProv;
                }
                int asvProvTotalInt = (int)Math.Floor(asvProvTotal);
                return asvProvTotalInt;
            }
            catch (Exception ex)
            {
                return $"Error converting rebar\n{ex.Message}";
            }
        }
        private void convertToAsv_Click(object sender, EventArgs e)
        {
            try
            {
                #region Get Excel Data
                Range sourceRange = ((RangeTextBox)attributeDic["beamAsvIn_BC"]).GetRangeForCurrentSheet();
                object[,] rebarProv = GetContentsAsObject2DArray(sourceRange);
                #endregion

                #region Conversion
                object[,] asProv = new object[rebarProv.GetLength(0), rebarProv.GetLength(1)];
                for (int rowNum = 0; rowNum < rebarProv.GetLength(0); rowNum++)
                {
                    for (int colNum = 0; colNum < rebarProv.GetLength(1); colNum++)
                    {
                        if (rebarProv[rowNum, colNum] != null)
                        {
                            string rebarStr = rebarProv[rowNum, colNum].ToString();
                            asProv[rowNum, colNum] = ConvertRebarStringToAsv(rebarStr);
                        }
                    }
                }
                #endregion

                #region Write to Excel
                Range writeRange = ((RangeTextBox)attributeDic["beamAsvOut_BC"]).GetRangeForCurrentSheet();
                WriteObjectToExcelRange(writeRange, 0, 0, true, asProv);
                #endregion
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Error"); }
        }

        #endregion
    }
}
