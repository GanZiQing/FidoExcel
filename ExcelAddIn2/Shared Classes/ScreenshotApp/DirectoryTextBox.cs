using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Button = System.Windows.Forms.Button;
using TextBox = System.Windows.Forms.TextBox;
using System.IO;

namespace ScreenshotApp
{
    public class DirectoryTextBox
    {
        #region Initialisation
        public string attName;
        public Button setButton;
        public Button openButton;
        public TextBox textBox;
        public string directory = "";

        public DirectoryTextBox(string attName, TextBox textBox, Button setButton, Button openButton = null)
        {
            this.attName = attName;
            this.setButton = setButton;
            if (openButton != null)
            {
                this.openButton = openButton;
            }
            this.textBox = textBox;
            SubscribeToEvents();
        }

        private void SubscribeToEvents()
        {
            setButton.Click += new EventHandler(setButton_Click);

            if (openButton != null)
            {
                openButton.Click += new EventHandler(openButton_Click);
            }

            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        #endregion

        #region Subscribed Events
        private void setButton_Click(object sender, EventArgs e)
        {
            CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
            if (customFolderBrowser.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            SetValue(customFolderBrowser.GetFolderPath());
            RefreshTextBox();
        }

        public void openButton_Click(object sender, EventArgs e)
        {
            // Check if button has been added
            if (openButton == null)
            {
                throw new Exception("Open Button not added to this directory");
            }

            // Get Path
            string folderPath = textBox.Text;

            //Check if path exist
            if (folderPath == "")
            {
                MessageBox.Show("No path provided", "Error");
                return;
            }

            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show($"Directory provided is invalid: \n\n{folderPath}?", "Error");
                return;
            }
            else
            {
                System.Diagnostics.Process.Start(folderPath);
            }
        }

        private void textBox_LostFocus(object sender, EventArgs e)
        {
            SetValueFromTextBox();
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        #endregion
        #region Set Value
        private void SetValue(string value)
        {
            directory = value;
        }

        public void SetValueFromTextBox()
        {
            if (textBox.Text == "")
            {
                directory = "";
                return;
            }

            if (!Path.IsPathRooted(textBox.Text) || !Directory.Exists(textBox.Text))
            {
                MessageBox.Show("Please provide a valid path", "Error");
                RefreshTextBox();
                return;
            }

            SetValue(textBox.Text);
        }

        private void RefreshTextBox()
        {
            textBox.Text = directory;
        }
        #endregion

        #region Check and Get Value
        public string CheckAndGetValue()
        {
            if (textBox.Text == "")
            {
                throw new Exception ($"Folder path not set");
            }
            else if (!Directory.Exists(textBox.Text))
            {
                string msg = $"Invalid folder path  provided:\n\n{textBox.Text}";
                throw new Exception(msg);
            }
            return textBox.Text;
        }
        #endregion
    }

    public class FileNameTextBox
    {
        #region Initialisation
        public TextBox textBox;
        public string attName;
        public string fileName = "";
        public string defaultExtension;
        public string defaultFileName;

        public FileNameTextBox(string attName, TextBox textBox, string defaultExtension = null, string defaultFileName = null)
        {
            this.attName = attName;
            this.textBox = textBox;
            this.defaultExtension = defaultExtension;
            this.defaultFileName = defaultFileName;
            textBox.Text = defaultFileName;
            SetValueFromTextBox();
            SubscribeToEvents();
        }

        private void SubscribeToEvents()
        {
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }
        #endregion

        #region Subscribed Events
        private void textBox_LostFocus(object sender, EventArgs e)
        {
            SetValueFromTextBox();
        }

        private void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.SelectNextControl((Control)sender, true, true, true, true);
            }
        }

        #endregion
        #region Set Value
        private void SetInternalValue(string value)
        {
            fileName = value;
        }

        private void SetValueFromTextBox(bool showMsg = false)
        {
            if (textBox.Text == "")
            {
                RefreshTextBox();
                return;
            }

            try
            {
                CheckIsFileName();
                if (defaultExtension != null)
                {
                    AppendExtension();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                RefreshTextBox();
                return;
            }

            SetInternalValue(textBox.Text);
            RefreshTextBox();
        }

        private void RefreshTextBox()
        {
            textBox.Text = fileName;
        }
        #endregion

        #region Check and Get Value
        private void CheckIsFileName()
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidFileNameChars());
            bool success = true;
            string invalidCharsInString = "";
            foreach (char c in textBox.Text)
            {
                if (invalidChars.Contains(c))
                {
                    invalidCharsInString += c + " ";
                    success = false;
                }
            }

            if (!success)
            {
                throw new Exception(($"'{textBox.Text}' contains invalid characters: \n" + invalidCharsInString));
            }
            return; 
        }

        private void AppendExtension()
        {
            if (defaultExtension == null)
            {
                throw new Exception("No extension defined");
            }

            if (Path.HasExtension(textBox.Text) && Path.GetExtension(textBox.Text) == defaultExtension)
            {
                return;
            }

            textBox.Text += defaultExtension;
            SetInternalValue(textBox.Text);
        }

        //public string CheckAndGetValue()
        //{
        //    if (textBox.Text == "")
        //    {
        //        throw new Exception($"Folder path not set");
        //    }
        //    else if (!Directory.Exists(textBox.Text))
        //    {
        //        string msg = $"Invalid folder path  provided:\n\n{textBox.Text}";
        //        throw new Exception(msg);
        //    }
        //    return textBox.Text;
        //}
        #endregion
    }

    class CustomFolderBrowser
    {
        OpenFileDialog dialog = new OpenFileDialog();
        public CustomFolderBrowser()
        {
            dialog.ValidateNames = false;  // Allows selecting folders
            dialog.Filter = "Folders|*. ";
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "Select Folder";  // Fake name to allow folder selection
        }

        public string folderPath = null;
        public DialogResult ShowDialog()
        {
            DialogResult dialogResult = dialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                string test = dialog.FileName;
                folderPath = Path.GetDirectoryName(dialog.FileName);
            }
            return dialogResult;
        }

        public string GetFolderPath()
        {
            if (folderPath == null)
            {
                throw new Exception("Folder path is not set");
            }
            return folderPath;
        }
    }
}


