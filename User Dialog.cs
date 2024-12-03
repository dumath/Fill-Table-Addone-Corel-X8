using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Fill_Table
{
    public partial class UserDialog : Form
    {
        private bool isValidatedStartNumber = false;
        private bool isValidatedWidth = false;
        private bool isValidatedHeight = false;

        internal int startNumber;
        internal double width;
        internal double height;

        public UserDialog()
        {
            InitializeComponent();
        }

        private void UserDialog_Load(object sender, EventArgs e) { }

        public void Calculate_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void StartNum_TextChanged(object sender, EventArgs e)
        {
            string s = (sender as TextBox).Text;
            
            if (int.TryParse(s, out startNumber))
            {
                isValidatedStartNumber = true;
                buttonSetActive();
            }
            else
            {
                isValidatedStartNumber = false;
                buttonSetActive();
                return;
            }
        }

        private void WidthNum_TextChanged(object sender, EventArgs e)
        {
            string s = (sender as TextBox).Text;

            if (double.TryParse(s, out width) && width > 0d)
            {
                isValidatedWidth = true;
                buttonSetActive();
            }
            else
            {
                isValidatedWidth = false;
                buttonSetActive();
                return;
            }
        }

        private void HeightNum_TextChanged(object sender, EventArgs e)
        {
            string s = (sender as TextBox).Text;

            if (double.TryParse(s, out height) && height > 0d)
            {
                isValidatedHeight = true;
                buttonSetActive();
            }
            else
            {
                isValidatedHeight = false;
                buttonSetActive();
                return;
            }
        }

        private void buttonSetActive()
        {
            if(isValidatedStartNumber)
            {
                if(isValidatedWidth)
                {
                    if(isValidatedHeight)
                    {
                        calculateButton.Enabled = true;
                        return;
                    }
                }
            }

            calculateButton.Enabled = false;
        }
    }
}
