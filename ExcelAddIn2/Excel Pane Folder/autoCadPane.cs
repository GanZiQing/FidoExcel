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
//using ACAS = Autodesk.AutoCAD.ApplicationServices;
//using ACDS = Autodesk.AutoCAD.DatabaseServices;
//using ADRT = Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Exception = System.Exception;

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
            #region Line Functions
            AddDynamicActionToButton(getLineCoords, () => InsertDynamicHeader(lineCoordinateHeaderGenerator));
            #endregion
        }

        #region Dynamic Header Generators
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
            CustomAttribute att;
            AttributeTextBox tbAtt;

            #region Line Functions
            #region Coordinates
            att = new CheckBoxAttribute("lineXCheck_AC", printXCheck, true);
            attDic.Add(att.attName, att);

            att = new CheckBoxAttribute("lineYCheck_AC", printYCheck, true);
            attDic.Add(att.attName, att);

            att = new CheckBoxAttribute("lineZCheck_AC", printZCheck);
            attDic.Add(att.attName, att);

            att = new CheckBoxAttribute("lineStartCheck_AC", printStartCheck, true);
            attDic.Add(att.attName, att);

            att = new CheckBoxAttribute("lineMidCheck_AC", printMidCheck, true);
            attDic.Add(att.attName, att);

            att = new CheckBoxAttribute("lineEndCheck_AC", printEndCheck, true);
            attDic.Add(att.attName, att);
            #endregion

            #region Properties
            att = new ComboBoxAttribute("lineProperty_AC", dispLineProperties, "Length");
            attDic.Add(att.attName, att);
            #endregion
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
            try
            {
                AdskGreeting();
            }
            catch (Exception ex) { MessageBox.Show($"Unable to get line coordinates\n{ex.Message}","Error"); }
        }

        public void AdskGreeting()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Starts a new transaction with the Transaction Manager
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                /* Creates a new MText object and assigns it a location,
                text value and text style */
                using (MText objText = new MText())
                {
                    // Specify the insertion point of the MText object
                    objText.Location = new Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0);

                    // Set the text string for the MText object
                    objText.Contents = "Greetings, Welcome to AutoCAD .NET";

                    // Set the text style for the MText object
                    objText.TextStyleId = acCurDb.Textstyle;

                    // Appends the new MText object to model space
                    acBlkTblRec.AppendEntity(objText);

                    // Appends to new MText object to the active transaction
                    acTrans.AddNewlyCreatedDBObject(objText, true);
                }

                // Saves the changes to the database and closes the transaction
                acTrans.Commit();
            }
        }
    }
}
