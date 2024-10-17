using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Properties = ExcelAddIn2.Properties;

namespace ScreenshotApp
{
    public partial class ScreenshotForm : Form
    {
        #region Initialise
        SettingsForm settingsForm;
        int[] dimensionValues;
        string[] dimensionLabels = new string[] { "width", "height", "X Position", "Y Position"};
        string latestFilePath;

        public ScreenshotForm(bool isBoundsForm = false)
        {
            InitializeComponent();
            SubscribeToEvents();
            InitializeDimensions();

            if (isBoundsForm)
            {
                takeScreenshot.Visible = false;
                openFolder.Visible = false;
                openFile.Visible = false;
                Settings.Visible = false;
                dispStatus.Visible = false;
                TopMost = false;
            }
            else
            {
                InitialiseSettings();
                dispStatus.Text = "";
                RegisterGlobalHotKey(Keys.A);
            }
            //StartPosition = FormStartPosition.CenterScreen;
            //Capture = true;
            //Cursor = Cursors.Default;
            //Activate();   
        }
        
        private void SubscribeToEvents()
        {
            Resize += new EventHandler((sender, e) => UpdatePositionInfo());
            Move += new EventHandler((sender, e) => UpdatePositionInfo());
            KeyPreview = true;

            FormClosing += new FormClosingEventHandler((sender, e) => UnregisterHotKey(this.Handle, 0));

            //ResizeBegin += new EventHandler(ShowMagnifier2);
            //ResizeEnd += new EventHandler(CloseMagnifier2);
            //KeyDown += new KeyEventHandler(ShowMagnifierKD);

            //ResizeBegin += new EventHandler(ShowMagnifier);
            //ResizeEnd += new EventHandler(CloseMagnifier);
        }

        private void InitialiseSettings()
        {
            settingsForm = new SettingsForm();
            string defaultFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Screenshots");
            if (!Directory.Exists(defaultFolderPath))
            {
                Directory.CreateDirectory(defaultFolderPath);
            }
            settingsForm.SetDirectory(defaultFolderPath);
            #region Subscribing
            settingsForm.AlwaysOnTopCheck.CheckedChanged += new EventHandler((sender, e) => ToggleTopState());
            settingsForm.HotKeyForm.HotKeySet += new EventHandler((sender, e) => ChangeHotKey());
            settingsForm.AspectRatioCheck.CheckedChanged += new EventHandler((sender, e) => AspectRatioSubscriptionHandler());
            #endregion

            //settingsForm.
        }

        private void ToggleTopState()
        {
            TopMost = settingsForm.AlwaysOnTopCheck.Checked;
            settingsForm.TopMost = settingsForm.AlwaysOnTopCheck.Checked;
        }

        #region Magnifier
        //Magnifier magnifier;
        //private EventHandler resizeEventHandler;

        //private void ShowMagnifier(object sender, EventArgs e)
        //{

        //    if (ModifierKeys == Keys.Control)
        //    {
        //        magnifier = new Magnifier();
        //        resizeEventHandler = new EventHandler((sender2, e2) => { MoveMagnifier(sender, null); });
        //        Resize += resizeEventHandler;
        //        magnifier.Show();
        //        MoveMagnifier(null, null);
        //    }
        //}

        //private void CloseMagnifier(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        magnifier.Close();
        //    }
        //    catch { }
        //    // Unsubscribe from the Resize event
        //    if (resizeEventHandler != null)
        //    {
        //        Resize -= resizeEventHandler;
        //        resizeEventHandler = null; // Clean up the handler reference
        //    }
        //}

        //private void MoveMagnifier(object sender, MouseEventArgs e)
        //{
        //    int dx = magnifier.Width / 2;
        //    int dy = magnifier.Height / 2;


        //    magnifier.Location = new Point(MousePosition.X - dx, MousePosition.Y - dy);
        //    dispStatus.Text = $"Mouse Position: {MousePosition.X}, {MousePosition.Y}\nFormPosition: {magnifier.Left},{magnifier.Top}";
        //}
        #endregion

        #endregion

        #region Dimensions
        TextBox[] dimensionTextBoxes;
        private void InitializeDimensions()
        {
            dimensionValues = new int[4];
            dimensionTextBoxes = new TextBox[] { dispWidth, dispHeight, dispXPosition, dispYPosition };
            UpdatePositionInfo();

            foreach(TextBox textBox in dimensionTextBoxes)
            {
                textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
                textBox.LostFocus += new EventHandler((sender, e) => SetDimensionFromInput(sender));
            }
        }

        private void UpdatePositionInfo()
        {
            dimensionValues = GetBounds();

            #region Show Values
            for (int i = 0; i < dimensionValues.Length; i++)
            {
                dimensionTextBoxes[i].Text = dimensionValues[i].ToString();
            }
            #endregion           
        }

        public int[] GetBounds()
        {
            int xOffset = (Bounds.Width - DisplayRectangle.Width) / 2;
            int xLoc = Left + xOffset;
            int yLoc = Top;
            int width = Width - 2 * xOffset;
            int height = Height - xOffset;
            // We use xOffset because yOffset includes top border. Bottom border is the same width as xOffset
            return new int[] { width, height, xLoc, yLoc};
        }

        private void SetBoundsFromSingleTextBox(int textBoxIndex, TextBox textBox)
        {
            #region Parse TextBox Value
            int newDimensionValue;
            try
            {
                newDimensionValue = int.Parse(textBox.Text);
            }
            catch //(Exception ex)
            {
                System.Media.SystemSounds.Asterisk.Play();
                dispStatus.Text = $"Unable to parse \"{textBox.Text}\" into integer for {dimensionLabels[textBoxIndex]}, value reset.";
                textBox.Text = dimensionValues[textBoxIndex].ToString();
                return;
            }
            #endregion

            #region Check if value needs to be reset
            dispStatus.Text = "";
            if (dimensionValues[textBoxIndex] == newDimensionValue) { return; }
            #endregion

            #region Set contorl dimension
            int xOffset = (Bounds.Width - DisplayRectangle.Width) / 2;
            switch (textBoxIndex)
            {
                case 0: // width
                    {
                        Width = newDimensionValue + 2 * xOffset;
                        break;
                    }
                case 1: // height
                    {
                        Height = newDimensionValue + 1 * xOffset;
                        break;
                    }
                case 2: // left
                    {
                        Left = newDimensionValue - xOffset;
                        break;
                    }
                case 3: // Top
                    {
                        Top = newDimensionValue;
                        break;
                    }
            }
            #endregion
        }
        
        private void SetDimensionFromInput(object sender)
        {
            TextBox textBox = (TextBox)sender;
            #region Get TextBox Index
            int textBoxIndex = -1;
            for (int i = 0; i < dimensionTextBoxes.Length; i++)
            {
                if (textBox == dimensionTextBoxes[i])
                {
                    textBoxIndex = i;
                    break;
                }
            }

            if (textBoxIndex == -1)
            {
                throw new Exception("Text box cannot be found.");
            }
            #endregion

            #region Set Value
            SetBoundsFromSingleTextBox(textBoxIndex, textBox);
            
            #endregion

        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        public void SetAllBounds(int[] dimensions)
        {
            int xOffset = (Bounds.Width - DisplayRectangle.Width) / 2;
            for (int textBoxIndex = 0; textBoxIndex < 4; textBoxIndex++)
            {
                switch (textBoxIndex)
                {
                    case 0: // width
                        {
                            Width = dimensions[textBoxIndex] + 2 * xOffset;
                            break;
                        }
                    case 1: // height
                        {
                            Height = dimensions[textBoxIndex] + 1 * xOffset;
                            break;
                        }
                    case 2: // left
                        {
                            Left = dimensions[textBoxIndex] - xOffset;
                            break;
                        }
                    case 3: // Top
                        {
                            Top = dimensions[textBoxIndex];
                            break;
                        }
                }
            }
        }
        #endregion

        #region Take Screenshot
        public int[] GetAdjustedDimensions()
        {
            // Adjust Dimensions in cases where scaling is done
            int[] adjustedValues = new int[4];
            dimensionValues.CopyTo(adjustedValues,0);

            #region Find Scale
            Screen currentScreen = Screen.FromControl(this);
            DeviceInfo screenInfo = ScreenHelper.GetTargetMonitorInfo(currentScreen.DeviceName);
            int physicalRes = screenInfo.HorizontalResolution;
            int virtualRes = screenInfo.MonitorArea.Width;
            float localScale = Convert.ToSingle(physicalRes) / Convert.ToSingle(virtualRes);
            #endregion

            #region Scale Width and Height
            adjustedValues[0] = Convert.ToInt32(Convert.ToSingle(dimensionValues[0]) * localScale);
            adjustedValues[1] = Convert.ToInt32(Convert.ToSingle(dimensionValues[1]) * localScale);
            #endregion

            #region Scale Coordinates
            // Coordinates are scaled according to the 'global scale' (min of all scales), then by the 'local scale' for anything within that screen
            // (Virtual) Global width = Physical width * local scale 
            // (Virtual) Global width = Local width / global scale * local scale
            // Local width = Physical Width / global scale

            // Local Width = MonitorArea.Width
            // Global Widht = cooridnate of the right edge, start of adjacent screen
            // Physical/Actual Width = Actual monitor resolution before any scaling

            // To convert coordinate to actual physical coordinates required by copy from screen:
            // Get the corner coordinate. Find the "actual" physical coordinate of the corner [CornerCoord x global scale]
            // Get the of the current coordinate relative to the corner. Find the "actual" physical coordinate offset [offset x local scale]
            // Super impose the two together [Corner + Offset]

            float globalScale = ScreenHelper.GetMinScale();
            Rectangle screenBounds = currentScreen.Bounds;
            float xLocal = dimensionValues[2] - screenBounds.X;
            float yLocal = dimensionValues[3] - screenBounds.Y;
            xLocal = Convert.ToInt32(xLocal * localScale);
            yLocal = Convert.ToInt32(yLocal * localScale);
            float xCorner = screenBounds.X * globalScale;
            float yCorner = screenBounds.Y * globalScale;

            adjustedValues[2] = Convert.ToInt32(xCorner + xLocal);
            adjustedValues[3] = Convert.ToInt32(yCorner + yLocal);
            #endregion

            //#region Test values
            //Rectangle rect2 = Bounds;
            //Screen[] allScreens = Screen.AllScreens;
            //Screen screen = Screen.FromControl(this);
            //Rectangle rect = Screen.GetBounds(this);
            //var screensInfo = ScreenHelper.GetMonitorsInfo();
            //#endregion

            return adjustedValues;
        }

        private void takeScreenshot_Click(object sender, EventArgs e)
        {
            try
            {
                int[] adjustedValues = GetAdjustedDimensions();
                int width = adjustedValues[0];
                int height = adjustedValues[1];
                int xLoc = adjustedValues[2];
                int yLoc = adjustedValues[3];

                Visible = false;
                Bitmap captureBitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                Graphics captureGraphics = Graphics.FromImage(captureBitmap);
                captureGraphics.CopyFromScreen(xLoc, yLoc, 0, 0, new Size(width, height));
                Visible = true;

                // Save image
                if (settingsForm.SaveFileChecked)
                {
                    string filePath = settingsForm.GetValidFilePath();
                    captureBitmap.Save(filePath, ImageFormat.Png);
                    latestFilePath = filePath;
                    dispStatus.Text = $"Screenshot saved as {Path.GetFileName(filePath)}";
                }

                // Copy image to clipboard
                if (settingsForm.ClipboardChecked)
                {
                    Clipboard.SetImage(captureBitmap);
                    if (!settingsForm.SaveFileChecked)
                    {
                        dispStatus.Text = $"Screenshot copied to clipboard.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void takeScreenshot_ClickObselete(object sender, EventArgs e)
        {
            try
            {
                int width = dimensionValues[0];
                int height = dimensionValues[1];
                int xLoc = dimensionValues[2];
                int yLoc = dimensionValues[3];

                Visible = false;
                Bitmap captureBitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                Graphics captureGraphics = Graphics.FromImage(captureBitmap);
                captureGraphics.CopyFromScreen(xLoc, yLoc, 0, 0, new Size(width, height));
                Visible = true;

                // Save image
                if (settingsForm.SaveFileChecked)
                {
                    string filePath = settingsForm.GetValidFilePath();
                    captureBitmap.Save(filePath, ImageFormat.Png);
                    latestFilePath = filePath;
                    dispStatus.Text = $"Screenshot saved as {Path.GetFileName(filePath)}";
                }

                // Copy image to clipboard
                if (settingsForm.ClipboardChecked)
                {
                    Clipboard.SetImage(captureBitmap);
                    if (!settingsForm.SaveFileChecked)
                    {
                        dispStatus.Text = $"Screenshot copied to clipboard.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        #endregion

        #region Settings and Open Buttons
        private void Settings_Click(object sender, EventArgs e)
        {
            try
            {
                Point settingsLocation = Location;
                settingsLocation.Offset(20, 150);
                settingsForm.Location = settingsLocation;
                settingsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void openFolder_Click(object sender, EventArgs e)
        {
            try
            {
                settingsForm.OpenFolder();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
            
        }

        private void openFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (latestFilePath == null)
                {
                    throw new Exception("No files found.");
                }
                System.Diagnostics.Process.Start(latestFilePath);
                //settingsForm.openFile_Click(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        public Button CloseButton
        {
            get { return closeButton; }
        }
        #endregion

        #region Keyboard Shortcuts

        #region Windows Constants
        // Constants for the modifier keys
        private const int MOD_ALT = 0x1;
        private const int MOD_CONTROL = 0x2;
        private const int MOD_SHIFT = 0x4;
        private const int MOD_WIN = 0x8;

        // Windows message ID for hotkey pressed
        private const int WM_HOTKEY = 0x0312;
        #endregion

        #region Register and Unregister HotKey
        // Import RegisterHotKey and UnregisterHotKey functions from user32.dll
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private void RegisterGlobalHotKey(Keys parsedKey)
        {
            int id = 0; // ID for the hotkey, can be any number
            RegisterHotKey(this.Handle, id, (uint)(MOD_CONTROL | MOD_SHIFT), (uint)parsedKey);
        }
        #endregion
        
        #region Change Hotkey
        private void ChangeHotKey()
        {
            Keys newKey = settingsForm.HotKeyForm.shortcutKey;
            UnregisterHotKey(Handle, 0);
            RegisterGlobalHotKey(newKey);
            dispStatus.Text = $"Keyboard shortcut changed to Ctrl + Shift + {newKey.ToString()}";
        }
        #endregion

        #region Hotkey Action
        // Override the WndProc method to capture the hotkey event
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                // Check which hotkey was pressed (if you have multiple)
                int id = m.WParam.ToInt32();
                if (id == 0) // The ID we assigned
                {
                    // Execute the code you want to run when the hotkey is pressed
                    PlayScreenshotSound();
                    takeScreenshot_Click(null, null);
                }
            }
            base.WndProc(ref m);
        }
        private void PlayScreenshotSound()
        {
            using (UnmanagedMemoryStream soundStream = Properties.Resources.Windows_Print_complete)
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(soundStream);
                player.Play();
            }
        }
        #endregion
        #endregion

        #region Aspect Ratio
        int aspectRatioWidth;
        int aspectRatioHeight;
        private void AspectRatioSubscriptionHandler()
        {
            bool checkState = settingsForm.AspectRatioCheck.Checked;
            if (checkState)
            {
                aspectRatioWidth = dimensionValues[0];
                aspectRatioHeight = dimensionValues[1];
                SizeChanged += FixAspectRatio;
            }
            else
            {
                SizeChanged -= FixAspectRatio;
            }
        }

        private void FixAspectRatio(object sender, EventArgs e)
        {
            double adjustedHeight = Convert.ToDouble(DisplayRectangle.Width)/ Convert.ToDouble(aspectRatioWidth) * Convert.ToDouble(aspectRatioHeight);
            int xOffset = (Bounds.Width - DisplayRectangle.Width) / 2;
            Height = (int)adjustedHeight + 1 * xOffset;
        }

        #endregion

        private void showScreenInfo_Click(object sender, EventArgs e)
        {
            Rectangle rect2 = Bounds;
            Screen[] allScreens = Screen.AllScreens;
            Screen screen = Screen.FromControl(this);
            Rectangle rect = Screen.GetBounds(this);
            var screensInfo = ScreenHelper.GetMonitorsInfo();
            string msg2 = "";
            foreach (Screen thisScreen in allScreens)
            {
                DeviceInfo screenInfo = ScreenHelper.GetTargetMonitorInfo(thisScreen.DeviceName);
                int physicalRes = screenInfo.HorizontalResolution;
                int virtualRes = screenInfo.MonitorArea.Width;
                float scale = Convert.ToSingle(physicalRes) / Convert.ToSingle(virtualRes);
                msg2 += $"{thisScreen.DeviceName}\n" +
                    $"Physical Resolution: {physicalRes}\n" +
                    $"Virtual Resolution: {virtualRes}\n" +
                    $"Local scale: {scale}\n" +
                    $"X: {screenInfo.MonitorArea.X}, Y:{screenInfo.MonitorArea.Y}\n" +
                    $"Virtual Width: {screenInfo.MonitorArea.Width}, Virtual Height: {screenInfo.MonitorArea.Height}\n\n";
            }
            MessageBox.Show(msg2);
        }
    }
}
