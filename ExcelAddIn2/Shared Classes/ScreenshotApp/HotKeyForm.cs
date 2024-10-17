using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenshotApp
{
    public partial class HotKeyForm : Form
    {
        public string keyboardKey = "A";
        public Keys shortcutKey = Keys.A;
        public HotKeyForm()
        {
            TopMost = true;
            StartPosition = FormStartPosition.CenterParent;
            InitializeComponent();
        }

        public Button SetHotKey
        {
            get { return setHotKey; }
        }

        public event EventHandler HotKeySet;

        private void setHotKey_Click(object sender, EventArgs e)
        {
            dispKeyboardKey.Text = dispKeyboardKey.Text.ToUpper();
            bool canParse = Enum.TryParse(dispKeyboardKey.Text, out Keys parsedKey);
            if (canParse)
            {
                shortcutKey = parsedKey;
                keyboardKey = dispKeyboardKey.Text;
                HotKeySet.Invoke(this,e);
                // trigger an event for parent form
                Close();
            }
            else
            {
                dispKeyboardKey.Text = shortcutKey.ToString();
                MessageBox.Show("Unable to set hotkey");
            }
            
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
