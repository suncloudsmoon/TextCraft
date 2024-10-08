﻿namespace TextForge
{
    partial class PasswordPrompt
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PasswordPrompt));
            this.PasswordTextBox = new System.Windows.Forms.TextBox();
            this.PasswordLabel = new System.Windows.Forms.Label();
            this.OkButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // PasswordTextBox
            // 
            this.PasswordTextBox.AccessibleDescription = "Enter your password for the PDF file here.";
            this.PasswordTextBox.Location = new System.Drawing.Point(12, 36);
            this.PasswordTextBox.Name = "PasswordTextBox";
            this.PasswordTextBox.Size = new System.Drawing.Size(374, 26);
            this.PasswordTextBox.TabIndex = 0;
            this.PasswordTextBox.UseSystemPasswordChar = true;
            // 
            // PasswordLabel
            // 
            this.PasswordLabel.AccessibleDescription = "Your PDF file added to RAG Control is encrypted. Please enter your password for t" +
    "hat PDF file to add it to the RAG Control.";
            this.PasswordLabel.AccessibleName = "";
            this.PasswordLabel.AutoSize = true;
            this.PasswordLabel.Location = new System.Drawing.Point(13, 13);
            this.PasswordLabel.Name = "PasswordLabel";
            this.PasswordLabel.Size = new System.Drawing.Size(352, 20);
            this.PasswordLabel.TabIndex = 1;
            this.PasswordLabel.Text = "Enter password for unlocking encrypted PDF file:";
            // 
            // OkButton
            // 
            this.OkButton.AccessibleDescription = "Unlocks the PDF using the password entered in the input box.";
            this.OkButton.AutoSize = true;
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(297, 68);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(89, 32);
            this.OkButton.TabIndex = 2;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // PasswordPrompt
            // 
            this.AcceptButton = this.OkButton;
            this.AccessibleDescription = "Enter your password for the encrypted PDF file.";
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(398, 100);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.PasswordLabel);
            this.Controls.Add(this.PasswordTextBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "PasswordPrompt";
            this.Text = "Security";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox PasswordTextBox;
        private System.Windows.Forms.Label PasswordLabel;
        private System.Windows.Forms.Button OkButton;
    }
}