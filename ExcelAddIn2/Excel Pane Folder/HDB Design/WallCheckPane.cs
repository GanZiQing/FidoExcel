using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn2.Excel_Pane_Folder.HDB_Design
{
    public partial class WallCheckPane : UserControl
    {
        #region Init
        Dictionary<string, AttributeTextBox> TextBoxAttributeDic = new Dictionary<string, AttributeTextBox>();
        Dictionary<string, CustomAttribute> OtherAttributeDic = new Dictionary<string, CustomAttribute>();
        public WallCheckPane()
        {
            InitializeComponent();
            CreateAttributes();
            AddToolTips();
        }

        private void AddToolTips()
        {
            #region Design Rebar
            toolTip1.SetToolTip(overwriteRebarCheck,
                "If unchecked, initial check will be done based on current values in the output range\n" +
                "If checked, initial check will be done based on values matched from Rebar Table");
            #endregion
        }

        private void CreateAttributes()
        {
            #region Match Reinforcement
            AttributeTextBox attTB = new RangeTextBox("rebarTable_WD", dispRebarTable, setRebarTable, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("storeyTable_WD", dispStoreyTable, setStoreyTable, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("pierLabelRange_WD", dispPierLabelRange, setPierLabelRange, "column");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("matchStoreyRange_WD", dispMatchStoreyCol, setMatchStoreyCol, "cell");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("outputRange_WD", dispOutputCol, setOutputCol, "cell");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new RangeTextBox("statusRange_WD", dispStatusCol, setStatusCol, "cell");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            var att = new CheckBoxAttribute("overwriteInitialRebar_WD", overwriteRebarCheck);
            att = new CheckBoxAttribute("overwriteInitialRebar_WD", unifyChangesCheck);
            #endregion

            #region Modify Reinforcement
            attTB = new RangeTextBox("mainRebarHeirarchy_WD", dispRebarHeirarchy, setRebarHeirarchy, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("targetUR_WD", dispTargetUR, true);
            attTB.type = "double";
            attTB.SetDefaultValue("0.9");
            TextBoxAttributeDic.Add(attTB.attName, attTB);

            attTB = new AttributeTextBox("maxAs_WD", dispMaxAs, true);
            attTB.type = "double";
            attTB.SetDefaultValue("4");
            TextBoxAttributeDic.Add(attTB.attName, attTB);


            #endregion

            #region Additional Settings
            att = new CheckBoxAttribute("backupSheetCheck_WD", backupSheetCheck);
            att = new CheckBoxAttribute("resetFontColorRebarTable_WD", resetFontColourRebarTableCheck);
            att = new CheckBoxAttribute("resetFontColorCheckSheet_WD", resetFontColourCheckSheetCheck);
            #endregion

            #region Decomposer
            attTB = new RangeTextBox("decomposeRange_WD", dispDecomposeRange, setDecomposeRange, "range");
            TextBoxAttributeDic.Add(attTB.attName, attTB);
            #endregion
        }
        #endregion
    }
}
