using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelAddIn2.CommonUtilities;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ExplorerBar;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class AutoCadPane: UserControl
    {
        #region Init
        Dictionary<string, object> attDic = new Dictionary<string, object>();
        public AutoCadPane()
        {
            InitializeComponent();
            CreateAttributes();
            AddToolTips();
            AddHeaders();
        }

        private void AddHeaders()
        {
            List<string> headers = null;

            #region Dynamic Headers for Get Line Coordinates
            AddDynamicActionToButton(getLineCoords, ()=>InsertDynamicHeader(lineCoordinateHeaderGenerator));
            #endregion
        }
        #region Dynamic Header Generator
        private object lineCoordinateHeaderGenerator()
        {
            try
            {
                List<string> axisToPrint = new List<string>();
                if (printXCheck.Checked) { axisToPrint.Add("X"); }
                if (printYCheck.Checked) { axisToPrint.Add("Y"); }
                if (printZCheck.Checked) { axisToPrint.Add("Z"); }
                if (axisToPrint.Count == 0) { throw new Exception("No coordinates to print"); }

                List<string> positionsToPrint = new List<string>();
                if (printStartCheck.Checked) { positionsToPrint.Add("Start"); }
                if (printMidCheck.Checked) { positionsToPrint.Add("Mid"); }
                if (printEndCheck.Checked) { positionsToPrint.Add("End"); }
                if (positionsToPrint.Count == 0) { throw new Exception("No coordinates to print"); }


                string[,] headerText = new string[2, axisToPrint.Count() * positionsToPrint.Count()];

                {
                    int colNum = 0;
                    // Create Top Row
                    foreach (string position in positionsToPrint)
                    {
                        headerText[0, colNum] = position;
                        colNum += axisToPrint.Count();
                    }
                }

                // Create Bottom Row
                for (int colNum = 0; colNum < axisToPrint.Count() * positionsToPrint.Count(); colNum++)
                {
                    int axisNum = colNum % axisToPrint.Count();
                    headerText[1, colNum] = axisToPrint[axisNum];
                }

                return headerText;
            }
            catch (Exception ex) { throw new Exception($"Unable to generate header for line coordinate\n{ex.Message}"); }
        }
        #endregion


        private void CreateAttributes()
        {
            #region Line Functions
            var att = new ComboBoxAttribute("lineCoordType_AC", dispLineOptions, "1 Start, mid, end");
            attDic.Add(att.attName, att);
            #endregion
        }

        private void AddToolTips()
        {
            ToolTip toolTip = new ToolTip();
            //#region Region
            //toolTip.SetToolTip(getline,
            //    "If folder name is empty, print files will be saved in current excel file path.\n" +
            //    "If folder name is provided, files are saved in a folder at the current excel file path.");
            //#endregion
        }
        #endregion


        private void getLineStartEnd_Click(object sender, EventArgs e)
        {
            lineCoordinateHeaderGenerator();
        }
    }
}
