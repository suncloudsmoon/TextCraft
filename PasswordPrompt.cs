using System;
using System.Windows.Forms;

namespace TextForge
{
    public partial class PasswordPrompt : Form
    {
        public string Password { get { return this.PasswordTextBox.Text; } }

        public PasswordPrompt()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }
    }
}
