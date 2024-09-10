using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using OpenAI.Chat;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class Forge
    {
        // Public
        public static readonly SystemChatMessage SystemPrompt = new SystemChatMessage("You are an expert writing assistant and editor, specialized in enhancing the clarity, coherence, and impact of text within a document. You analyze text critically and provide constructive feedback to improve the overall quality.");

        // Private
        private AboutBox _box;
        private static RibbonGroup _optionsBox;

        private CustomTaskPane _generateTaskPane;
        private CustomTaskPane _ragControlTaskPane;

        private void Forge_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                if (!ThisAddIn.IsAddinInitialized)
                    ThisAddIn.InitializeAddin();

                List<string> models = new List<string>(ThisAddIn.ModelList);

                // Remove embedding models from the list
                var copyList = new List<string>(models);
                foreach (var model in copyList)
                {
                    foreach (var item in ThisAddIn.EmbedModels)
                        if (model.Contains(item))
                            models.Remove(model);
                    if (model.Contains("embed"))
                        models.Remove(model);
                }

                var ribbonFactory = Globals.Factory.GetRibbonFactory();
                foreach (string model in models)
                {
                    var newItem = ribbonFactory.CreateRibbonDropDownItem();
                    newItem.Label = model;

                    ModelListDropDown.Items.Add(newItem);

                    if (model == ThisAddIn.Model)
                    {
                        ModelListDropDown.SelectedItem = newItem;
                        UpdateCheckbox();
                    }
                }

                _box = new AboutBox();
                _optionsBox = this.OptionsGroup;

                _generateTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new GenerateUserControl(), this.GenerateButton.Label);
                _ragControlTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(ThisAddIn.RagControl, this.RAGControlButton.Label);
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void GenerateButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _generateTaskPane.Visible = !_generateTaskPane.Visible;
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void ModelListDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.Model = ModelListDropDown.SelectedItem.Label;
                ThisAddIn.ContextLength = RAGControl.GetContextLength(ThisAddIn.Model);
                UpdateCheckbox();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }
        private void DefaultCheckBox_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (this.DefaultCheckBox.Checked)
                    Properties.Settings.Default.DefaultModel = ModelListDropDown.SelectedItem.Label;
                else
                    Properties.Settings.Default.DefaultModel = null;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void RAGControlButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _ragControlTaskPane.Visible = !_ragControlTaskPane.Visible;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _box.ShowDialog();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }
        private void CancelButton_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                ThisAddIn.CancellationTokenSource.Cancel();
                CancelButtonVisibility(false);
                ThisAddIn.CancellationTokenSource = new CancellationTokenSource();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void WritingToolsGallery_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                switch (e.Control.Id)
                {
                    case "ReviewButton":
                        await ReviewButton_Click();
                        break;
                    case "ProofreadButton":
                        await ProofreadButton_Click();
                        break;
                    case "RewriteButton":
                        await RewriteButton_Click();
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static async Task ReviewButton_Click()
        {
            Word.Paragraphs paragraphs = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs;
            const string prompt = "As an expert writing assistant, suggest specific improvements to the paragraph, focusing on clarity, coherence, structure, grammar, and overall effectiveness. Ensure that your suggestions are detailed and aimed at improving the paragraph within the context of the entire Document.";

            bool hasCommented = false; 
            if (Globals.ThisAddIn.Application.Selection.End - Globals.ThisAddIn.Application.Selection.Start > 0)
            {
                await AddComment(Globals.ThisAddIn.Application.ActiveDocument.Comments, Globals.ThisAddIn.Application.Selection.Range, Review(paragraphs, Globals.ThisAddIn.Application.Selection.Range, prompt));
                hasCommented = true;
            }
            else
            {
                foreach (Word.Paragraph p in paragraphs)
                    // It isn't a paragraph if it doesn't contain a full stop.
                    if (p.Range.Text.Contains('.'))
                    {
                        await AddComment(Globals.ThisAddIn.Application.ActiveDocument.Comments, p.Range, Review(paragraphs, p.Range, prompt));
                        hasCommented = true;
                    }

            }
            if (!hasCommented)
                MessageBox.Show("Review complete!", "Action Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static async Task ProofreadButton_Click()
        {
            await AnalyzeText(
                "You are a proofreading assistant. Your task is to correct any spelling, grammar, punctuation, or stylistic errors in the provided text. Only make necessary changes directly in the text. If the text is already correct and does not require any changes, simply repeat the text as it is without providing any explanations or comments.",
                "Please proofread the following text. Make any necessary corrections directly in the text without adding any explanations or comments. If the text is correct and needs no changes, just repeat it exactly as it is:"
                );
        }

        private static async Task RewriteButton_Click()
        {
            await AnalyzeText(
                "You are an advanced language model tasked with rewriting text provided by the user. You should focus on enhancing clarity, improving flow, and maintaining the original meaning of the text. Your responses should strictly consist of the rewritten text with no additional explanations or comments.",
                "Please rewrite the following text:"
                );
        }

        private static async Task AnalyzeText(string systemPrompt, string userPrompt)
        {
            var selectionRange = Globals.ThisAddIn.Application.Selection.Range;
            var range = (selectionRange.End - selectionRange.Start > 0) ? selectionRange : throw new InvalidRangeException("No text is selected for analysis!");
            
            ChatClient client = new ChatClient(ThisAddIn.Model, ThisAddIn.ApiKey, ThisAddIn.ClientOptions);
            var streamingAnswer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>() { new SystemChatMessage(systemPrompt), new UserChatMessage(@$"{userPrompt}: {range.Text}") },
                new ChatCompletionOptions() { MaxTokens = ThisAddIn.ContextLength },
                ThisAddIn.CancellationTokenSource.Token
            );
            range.Delete();
            await AddStreamingContentToRange(streamingAnswer, range);
            Globals.ThisAddIn.Application.Selection.SetRange(range.Start, range.End);
        }

        public static async Task AddStreamingContentToRange(AsyncCollectionResult<StreamingChatCompletionUpdate> streamingAnswer, Word.Range range)
        {
            StringBuilder response = new StringBuilder();
            CancelButtonVisibility(true);
            await foreach (var update in streamingAnswer.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
            {
                if (ThisAddIn.CancellationTokenSource.IsCancellationRequested) break;
                foreach (var newContent in update.ContentUpdate)
                {
                    range.Text += newContent.Text;
                    response.Append(newContent.Text);
                }
            }
            CancelButtonVisibility(false);

            range.Text = ThisAddIn.RemoveMarkdownSyntax(response.ToString());
            ThisAddIn.ApplyAllMarkdownFormatting(range, response.ToString());
        }

        public static void CancelButtonVisibility(bool option)
        {
            _optionsBox.Visible = option;
        }

        private void UpdateCheckbox()
        {
            DefaultCheckBox.Checked = (Properties.Settings.Default.DefaultModel == ThisAddIn.Model);
        }

        public static async Task AddComment(Word.Comments comments, Word.Range range, AsyncCollectionResult<StreamingChatCompletionUpdate> streamingContent)
        {
            Word.Comment c = comments.Add(range, string.Empty);
            c.Author = ThisAddIn.Model;
            Word.Range commentRange = c.Range.Duplicate; // Duplicate the range to work with

            StringBuilder comment = new StringBuilder();
            // Run the comment generation in a background thread
            await Task.Run(async () =>
            {
                await foreach (var update in streamingContent.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
                {
                    if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                        break;
                    foreach (var content in update.ContentUpdate)
                    {
                        commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd); // Move to the end of the range
                        commentRange.Text = content.Text; // Append new text
                        commentRange = c.Range.Duplicate; // Update the range to include the new text
                        comment.Append(content.Text);
                    }
                }
                c.Range.Text = ThisAddIn.RemoveMarkdownSyntax(comment.ToString());
            });
        }

        private static AsyncCollectionResult<StreamingChatCompletionUpdate> Review(Word.Paragraphs context, Word.Range p, string prompt)
        {
            var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            UserChatMessage commentPrompt = new UserChatMessage($@"Please review the following paragraph extracted from the Document: ""{RAGControl.SubstringTokens(p.Text, (int)(ThisAddIn.ContextLength * 0.2))}""{Environment.NewLine}{prompt}");
            return RAGControl.AskQuestion(SystemPrompt, new List<UserChatMessage> { commentPrompt }, docRange);
        }
    }
}
