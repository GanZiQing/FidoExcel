using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExcelAddIn2.Excel_Pane_Folder;
using System.Windows.Forms.Integration;

namespace ExcelAddIn2
{
    public partial class LauncherRib
    {
        private void Launcher_Load(object sender, RibbonUIEventArgs e)
        {

        }

        #region Helper Functions
        private List<CustomTaskPane> GetWindowPanes(ref List<CustomTaskPane> thisPaneList)
        {
            //Returns list of the window panes in the active window
            int numExistingPanes = 0;
            Window currentWindow = Globals.ThisAddIn.Application.ActiveWindow;
            List<CustomTaskPane> thisWindowPaneList = new List<CustomTaskPane>();
            List<CustomTaskPane> copyThisPaneList = thisPaneList.ToList();
            foreach (CustomTaskPane pane in copyThisPaneList)
            {
                try
                {
                    if (((Window)pane.Window).Caption != currentWindow.Caption) { continue; }
                    numExistingPanes += 1;
                    thisWindowPaneList.Add(pane);
                }
                catch (COMException ex) when (ex.Message.Contains("The taskpane has been deleted or is otherwise no longer valid"))
                {
                    thisPaneList.Remove(pane);
                }
            }
            return thisWindowPaneList;
        }

        private void AddPane<T>(ref List<CustomTaskPane> PaneList, string title) where T : UserControl, new()
        {
            T PaneControl = new T();
            if (PaneList.Count > 0) // Add new Panes to list 
            {
                title += " " + (PaneList.Count + 1).ToString();
            }
            #region Set Size
            int width = PaneControl.Width + SystemInformation.VerticalScrollBarWidth*2;
            CustomTaskPane PaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(PaneControl, title);
            PaneValue.Width = width;
            #endregion

            PaneValue.Visible = true;
            PaneList.Add(PaneValue);
        }

        private void TogglePaneVisibility(List<CustomTaskPane> PaneList, int MaxPanes)
        {
            // Count how many visible
            int numVisible = 0;
            Window currentWindow = Globals.ThisAddIn.Application.ActiveWindow;
            foreach (CustomTaskPane pane in PaneList)
            {
                if (pane.Visible)
                {
                    numVisible += 1;
                }
            }

            // Toggle visibility of one pane 
            if (numVisible < MaxPanes)
            {
                foreach (CustomTaskPane PaneValue in PaneList)
                {
                    if (!PaneValue.Visible)
                    {
                        PaneValue.Visible = true;
                        break;
                    }
                }
            }
            else
            {
                foreach (CustomTaskPane PaneValue in PaneList)
                {
                    PaneValue.Visible = false;
                }
            }
        }
        #endregion

        #region ETABS Pane Launcher
        private List<CustomTaskPane> ETABSPaneList = new List<CustomTaskPane>();
        private void ETABSPaneLauncher_Click(object sender, RibbonControlEventArgs e)
        {
            // User Inputs
            int NumPanes = 2;
            List<CustomTaskPane> thisPaneList = ETABSPaneList;
            string title = "ETABS";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<ETABSTaskPane>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }
        #endregion


        #region Iteration Pane Launcher
        private List<CustomTaskPane> IterationPaneList = new List<CustomTaskPane>();
        private void IterationTools_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 2;
            List<CustomTaskPane> thisPaneList = IterationPaneList;
            string title = "Iteration Tools";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<IterationPane>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }


        #endregion

        #region Print Tools Pane Launcher
        private List<CustomTaskPane> PrintPaneList = new List<CustomTaskPane>();
        private void PrintToolsButton_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 1;
            List<CustomTaskPane> thisPaneList = PrintPaneList;
            string title = "Directory and Pdf";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<DirectoryAndPdf>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }


        #endregion

        #region Format Tools Launcher
        private List<CustomTaskPane> formatToolsList = new List<CustomTaskPane>();
        private void TestButt_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 1;
            List<CustomTaskPane> thisPaneList = formatToolsList;
            string title = "Format Tools";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<FormatToolsPane>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }
        #endregion

        #region Graph Tools Launcher
        private List<CustomTaskPane> graphToolsList = new List<CustomTaskPane>();
        private void PlottingTools_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 1;
            List<CustomTaskPane> thisPaneList = graphToolsList;
            string title = "Graph Tools";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<GraphToolsPane>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }
        #endregion

        #region Piling Tools Launcher
        private List<CustomTaskPane> pilingToolsList = new List<CustomTaskPane>();
        //private void PilingToolsButton_Click(object sender, RibbonControlEventArgs e)
        //{
        //    int NumPanes = 1;
        //    List<CustomTaskPane> thisPaneList = pilingToolsList;
        //    string title = "Piling Tools";

        //    #region Default Code - Replace Task Pane Type
        //    List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

        //    if (windowTaskPane.Count < NumPanes) // add new panes to list 
        //    {
        //        AddPane<PilingPane>(ref thisPaneList, title);
        //    }
        //    else // Start toggling visibility of lists
        //    {
        //        TogglePaneVisibility(windowTaskPane, NumPanes);
        //    }
        //    #endregion
        //}
        #endregion

        #region Report Pane
        private List<CustomTaskPane> reportList = new List<CustomTaskPane>();
        private void reportPane_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 1;
            List<CustomTaskPane> thisPaneList = reportList;
            string title = "Report Tools";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<ReportPane>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }

        #endregion

        #region Beam Design
        private List<CustomTaskPane> beamList = new List<CustomTaskPane>();
        private void beamDesign_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 1;
            List<CustomTaskPane> thisPaneList = beamList;
            string title = "Beam Design Tools";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<BeamDesign>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }
        #endregion



        #region Wall Design
        private List<CustomTaskPane> wallList = new List<CustomTaskPane>();
        private void wallDesign_Click(object sender, RibbonControlEventArgs e)
        {
            int NumPanes = 1;
            List<CustomTaskPane> thisPaneList = wallList;
            string title = "Wall Design Tools";

            #region Default Code - Replace Task Pane Type
            List<CustomTaskPane> windowTaskPane = GetWindowPanes(ref thisPaneList);

            if (windowTaskPane.Count < NumPanes) // add new panes to list 
            {
                AddPane<WallDesign>(ref thisPaneList, title);
            }
            else // Start toggling visibility of lists
            {
                TogglePaneVisibility(windowTaskPane, NumPanes);
            }
            #endregion
        }
        #endregion
    }
}
