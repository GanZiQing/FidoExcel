using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn2.Excel_Pane_Folder
{
    public partial class basePane: UserControl
    {
        #region Init
        Dictionary<string, object> attributeDic = new Dictionary<string, object>();
        public basePane()
        {
            InitializeComponent();
            //CreateAttributes();
            //AddToolTips();
            //AddHeaders();
        }

        private void AddHeaders()
        {
            //List<string> headers = new List<string>
            //{
            //"Section Title",
            //"Start Pg Num",
            //"End Pg Num",
            //"Total Pg Num",
            //"Insert New Page",
            //"File Path",
            //};
            //AddHeaderMenuToButton(getLineStartEnd, headers);
        }

        private void CreateAttributes()
        {
            //var thisAtt = new ComboBoxAttribute("attName", getLineStartEnd, , "1 Start, mid, end");
            //attributeDic.Add(thisAtt.attName, thisAtt);

        }

        private void AddToolTips()
        {
            ToolTip toolTip = new ToolTip();
            //#region Region

            //toolTip.SetToolTip(button,
            //    "If folder name is empty, print files will be saved in current excel file path.\n" +
            //    "If folder name is provided, files are saved in a folder at the current excel file path.");
            //#endregion
        }
        #endregion
    }
}
