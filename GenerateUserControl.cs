using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenAI.Chat;

namespace TextForge
{
    public partial class GenerateUserControl : UserControl
    {
        private readonly SystemChatMessage _systemPrompt = new SystemChatMessage("You are an AI assistant designed to help users create content based on existing documents. Your task is to understand the user’s query and the context provided by the existing document, and then generate relevant and coherent content. Ensure that the content is accurate, well-structured, and aligns with the user’s requirements.");
        private readonly CultureLocalizationHelper _cultureHelper = new CultureLocalizationHelper("TextForge.GenerateUserControl", typeof(GenerateUserControl).Assembly);

        public GenerateUserControl()
        {
            try
            {
                InitializeComponent();
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void GenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                string textBoxContent = this.PromptTextBox.Text;
                if (textBoxContent.Length == 0)
                    throw new TextBoxEmptyException(_cultureHelper.GetLocalizedString("[GenerateButton_Click] TextBoxEmptyException #1"));
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
            catch (TextBoxEmptyException ex)
            {
                CommonUtils.DisplayInformation(ex);
            }
            catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void PromptTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control && e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    this.GenerateButton.PerformClick();
                }
                else if (e.Control && e.KeyCode == Keys.Back)
                {
                    e.SuppressKeyPress = true;
                    int cursorPosition = this.PromptTextBox.SelectionStart;
                    string text = this.PromptTextBox.Text;

                    // Handle multiple trailing spaces
                    while (cursorPosition > 0 && text[cursorPosition - 1] == ' ')
                    {
                        cursorPosition--;
                    }

                    text = text.TrimEnd();

                    if (string.IsNullOrWhiteSpace(text))
                    {
                        this.PromptTextBox.Clear();
                        this.PromptTextBox.SelectionStart = 0;
                    }
                    else
                    {
                        int lastSpaceIndex = text.LastIndexOf(' ', cursorPosition - 1);
                        if (lastSpaceIndex != -1)
                        {
                            // Retain a space after deletion
                            this.PromptTextBox.Text = text.Remove(lastSpaceIndex + 1, cursorPosition - lastSpaceIndex - 1);
                            this.PromptTextBox.SelectionStart = lastSpaceIndex + 1;
                        }
                        else
                        {
                            this.PromptTextBox.Text = text.Remove(0, cursorPosition);
                            this.PromptTextBox.SelectionStart = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }
    }

    public class TextBoxEmptyException : ArgumentException
    {
        public TextBoxEmptyException(string message) : base(message) { }
    }
}
