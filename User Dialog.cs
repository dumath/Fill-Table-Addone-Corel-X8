using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Create_Editable_Cells
{
    public partial class Preference : Form
    {
        private Main macro;

        public Preference(Main macro)
        {
            InitializeComponent();

            this.macro = macro;

            outlineWidthComboBox.DataSource = new double[] { 0.00, 0.05, 0.1, 0.2, 0.25, 0.5, 1 };
        }

        private void Preference_Load(object sender, EventArgs e) { }

        private void startNumberTextBox_TextChanged(object sender, EventArgs e)
        {
            string text = (sender as TextBox).Text;

            macro.RefreshStartNumber(text, previewCheckBox.Checked);
        }

        private void cellWidthTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox input = sender as TextBox;

            createMapButton.Enabled = macro.RefreshCellWidth(input.Text, previewCheckBox.Checked);
        }

        private void cellHeightTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox input = sender as TextBox;

            createMapButton.Enabled = macro.RefreshCellHeight(input.Text, previewCheckBox.Checked);
        }

        private void marginNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown input = sender as NumericUpDown;

            macro.RefreshMargin(input.Value, previewCheckBox.Checked);
        }
        
        private void OutlineWidth_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox input = sender as ComboBox;

            macro.RefreshOutline(input.SelectedValue.ToString(), previewCheckBox.Checked);
        }

        private void previewCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            macro.CreatePreviewMap();
        }

        private void createMapButton_Click(object sender, EventArgs e)
        {
            macro.CreateMap();
        }
    }
}
