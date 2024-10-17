using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn2
{
    public partial class CustomInputBox : Form
    {
        #region Init

        public bool setValue = false;
        public string type = "string";
        public CustomInputBox(string value = null)
        {
            InitializeComponent();
            if (value != null)
            {
                valueTextBox.Text = value;
            }
            CancelButton = cancelButton;
            StartPosition = FormStartPosition.CenterParent;
            SubscribeToEvents();
        }

        private void SubscribeToEvents()
        {
            valueTextBox.KeyDown += new KeyEventHandler(valueTextBox_KeyDown);
        }
        protected void valueTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                okButton_Click(sender, e);
            }
        }

        #endregion

        private void okButton_Click(object sender, EventArgs e)
        {
            bool inputOk = CheckInput();
            if (!inputOk)
            {
                MessageBox.Show("Invalid Input","Error");
                return;
            }
            setValue = true;
            Close();
        }

        #region Verification
        private bool CheckInput()
        {
            if (valueTextBox.Text == "")
            {
                return false;
            }

            switch (type)
            {
                case "string":
                    return true;
                case "int":
                    return Int32.TryParse(valueTextBox.Text, out _);
                case "double":
                    return double.TryParse(valueTextBox.Text, out _);
                default:
                    throw new Exception("Verification type not found");
            }
        }

        #endregion

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        #region Get and Set
        public string Value
        {
            get
            {
                if (!setValue)
                {
                    return null;
                }
                else
                {
                    return valueTextBox.Text;
                }
            }
            set
            {
                valueTextBox.Text = value;
            }
        }

        public string Instruction
        {
            set { instruction.Text = value; }
        }
        #endregion

    }
}
