using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenAI.Chat;

namespace TextForge
{
    public partial class GenerateUserControl : UserControl
    {
        private readonly SystemChatMessage _systemPrompt = new SystemChatMessage("You are an AI assistant designed to help users create content based on existing documents. Your task is to understand the user’s query and the context provided by the existing document, and then generate relevant and coherent content. Ensure that the content is accurate, well-structured, and aligns with the user’s requirements.");
        public GenerateUserControl()
        {
            try
            {
                InitializeComponent();
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void GenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                string textBoxContent = this.PromptTextBox.Text;
                if (textBoxContent.Length == 0)
                    throw new ArgumentException("The textbox is empty!");

                /*
                 * So, If the user changes the selection carot in Word after clicking "generate" (bc it takes so long to generate text).
                 * Then, it won't affect where the text is placed.
                 */
                var rangeBeforeChat = Globals.ThisAddIn.Application.Selection.Range;
                var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
                var streamingAnswer = RAGControl.AskQuestion(
                    _systemPrompt,
                    new List<UserChatMessage> { new UserChatMessage(textBoxContent) },
                    docRange
                );

                // Clear any selected text by the user
                if (rangeBeforeChat.End - rangeBeforeChat.Start > 0)
                    rangeBeforeChat.Delete();

                await Forge.AddStreamingContentToRange(streamingAnswer, rangeBeforeChat);
                Globals.ThisAddIn.Application.Selection.SetRange(rangeBeforeChat.Start, rangeBeforeChat.End);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PromptTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                this.GenerateButton.PerformClick();
            }
        }
    }
}
