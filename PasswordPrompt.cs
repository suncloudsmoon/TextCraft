using System;
using System.Windows.Forms;

namespace TextForge
{
    public partial class PasswordPrompt : Form
    {
        public string Password { get { return this.PasswordTextBox.Text; } }

        public PasswordPrompt()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
