namespace TextForge
{
    partial class GenerateUserControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GenerateUserControl));
            this.PromptTextBox = new System.Windows.Forms.TextBox();
            this.GenerateButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // PromptTextBox
            // 
            resources.ApplyResources(this.PromptTextBox, "PromptTextBox");
            this.PromptTextBox.Name = "PromptTextBox";
            this.PromptTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PromptTextBox_KeyDown);
            // 
            // GenerateButton
            // 
            resources.ApplyResources(this.GenerateButton, "GenerateButton");
            this.GenerateButton.Name = "GenerateButton";
            this.GenerateButton.UseVisualStyleBackColor = true;
            this.GenerateButton.Click += new System.EventHandler(this.GenerateButton_Click);
            // 
            // GenerateUserControl
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.GenerateButton);
            this.Controls.Add(this.PromptTextBox);
            this.Name = "GenerateUserControl";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox PromptTextBox;
        private System.Windows.Forms.Button GenerateButton;
    }
}
