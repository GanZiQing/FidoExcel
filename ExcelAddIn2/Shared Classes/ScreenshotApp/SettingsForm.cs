using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace ScreenshotApp
{
    public partial class SettingsForm : Form
    {
        DirectoryTextBox directoryTextBox;
        FileNameTextBox fileNameTextBox;
        private HotKeyForm hotKeyForm = new HotKeyForm();

        public SettingsForm()
        {
            InitializeComponent();
            directoryTextBox = new DirectoryTextBox("directory", dispFolderPath, setFolder, openFolder);
            fileNameTextBox = new FileNameTextBox("filename", dispFileName, ".png", "Screenshot");
            //StartPosition = FormStartPosition.CenterParent;
            StartPosition = FormStartPosition.Manual;
            CancelButton = cancelButton;
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; } // Prevents the form from taking focus
        }

        public void openFile_Click(object sender, EventArgs e)
        {
            string filePath = GetValidFilePath();
            try
            {
                System.Diagnostics.Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        public void OpenFolder()
        {
            directoryTextBox.openButton_Click(null, null);
        }

        public void SetDirectory(string directory)
        {
            directoryTextBox.textBox.Text = directory;
            directoryTextBox.SetValueFromTextBox();
        }
        #region Get Values
       public string GetValidFilePath()
        {
            // Check that directory is valid
            if (!Directory.Exists(directoryTextBox.directory))
            {
                throw new Exception($"Directory provided does not exist:\n{directoryTextBox.directory}");
            }

            // Get file name
            if (overwriteCheck.Checked)
            {
                return Path.Combine(directoryTextBox.directory, fileNameTextBox.fileName);
            }
            else
            {
                return GetAvailableFileName();
            }
        }

        private string GetAvailableFileName()
        {
            string filePath = Path.Combine(directoryTextBox.directory, fileNameTextBox.fileName);

            if (!File.Exists(filePath))
            {
                return filePath;
            }

            // Get the directory, filename, and extension
            string directory = Path.GetDirectoryName(filePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            int fileNumber = 2;

            // Continue looping until a file with the new name does not exist
            string newFilePath = filePath;
            while (File.Exists(newFilePath))
            {
                newFilePath = Path.Combine(directory, $"{fileNameWithoutExtension} ({fileNumber}){extension}");
                fileNumber++;
            }

            return newFilePath;
        }

        //public string FileName
        //{
        //    get { return fileNameTextBox.fileName; }
        //}
        //public string FolderPath
        //{
        //    get { return directoryTextBox.directory; }
        //}

        public bool ClipboardChecked
        {
            get { return clipboardCheck.Checked; }
            //set { clipboardCheck.Checked = value; }
        }

        public bool SaveFileChecked
        {
            get { return saveFileCheck.Checked; }
            //set { return saveFileCheck.Checked; }
        }
        #endregion

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        #region Exposing Internals
        public CheckBox AlwaysOnTopCheck
        {
            get { return alwaysOnTopCheck; }
        }

        public CheckBox AspectRatioCheck
        {
            get { return aspectRatioCheck; }
        }
        #endregion

        #region HotKey
        public HotKeyForm HotKeyForm
        {
            get { return hotKeyForm; }
        }
        private void ShowHotKeyForm_Click(object sender, EventArgs e)
        {
            hotKeyForm.ShowDialog();
        }
        #endregion
    }
}
