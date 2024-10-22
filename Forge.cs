﻿using System;
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
        public static readonly CultureLocalizationHelper CultureHelper = new CultureLocalizationHelper("TextForge.Forge", typeof(Forge).Assembly);
        public static readonly object InitializeDoor = new object();
        public static readonly SystemChatMessage CommentSystemPrompt = new SystemChatMessage("You are an expert writing assistant and editor, specialized in enhancing the clarity, coherence, and impact of text within a document. You analyze text critically and provide constructive feedback to improve the overall quality.");

        // Private
        private AboutBox _box;
        private static RibbonGroup _optionsBox;

        private void Forge_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                    ThisAddIn.AddTaskPanes(Globals.ThisAddIn.Application.ActiveDocument);

                Thread startup = new Thread(InitializeForge);
                startup.SetApartmentState(ApartmentState.STA);
                startup.Start();
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InitializeForge() {
            try
            {
                lock (InitializeDoor)
                {
                    if (!ThisAddIn.IsAddinInitialized)
                        ThisAddIn.InitializeAddIn();

                    List<string> modelList = new List<string>(ModelProperties.GetModelList(ThisAddIn.ModelList));

                    // Remove embedding models from the list
                    modelList = RemoveEmbeddingModels(modelList).ToList();
                    AddEmbeddingModelsToDropDownList(modelList);
                }
                _box = new AboutBox();
                _optionsBox = this.OptionsGroup;
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private IEnumerable<string> RemoveEmbeddingModels(IEnumerable<string> modelList)
        {
            return modelList
                .Where(model => !ModelProperties.UniqueEmbedModels.Any(item => model.Contains(item)) && !model.Contains("embed"))
                .ToList();
        }

        private void AddEmbeddingModelsToDropDownList(IEnumerable<string> models)
        {
            var ribbonFactory = Globals.Factory.GetRibbonFactory();
            var sortedModels = models.OrderBy(m => m).ToList();
            foreach (string model in sortedModels)
            {
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
            }
        }

        private async void ModelListDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                await Task.Run(() =>
                {
                    ThisAddIn.Model = GetSelectedItemLabel();
                    UpdateCheckbox();
                    ThisAddIn.ContextLength = ModelProperties.GetContextLength(ThisAddIn.Model, ThisAddIn.ModelList); // this request is slow
                });
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
                    Properties.Settings.Default.DefaultModel = GetSelectedItemLabel();
                else
                    Properties.Settings.Default.DefaultModel = null;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private string GetSelectedItemLabel()
        {
            return ModelListDropDown.SelectedItem.Label;
        }

        private void GenerateButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var taskpanes = ThisAddIn.AllTaskPanes[Globals.ThisAddIn.Application.ActiveDocument];
                taskpanes.Item1.Visible = !taskpanes.Item1.Visible;
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
                var taskpanes = ThisAddIn.AllTaskPanes[Globals.ThisAddIn.Application.ActiveDocument];
                taskpanes.Item2.Visible = !taskpanes.Item2.Visible;
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
                CancelButtonVisibility(false);
                ThisAddIn.CancellationTokenSource.Cancel();
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
                        throw new ArgumentOutOfRangeException(CultureHelper.GetLocalizedString("[WritingToolsGallery_ButtonClick] ArgumentOutOfRangeException #1"));
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static async Task ReviewButton_Click()
        {
            const string prompt = "As an expert writing assistant, suggest specific improvements to the paragraph, focusing on clarity, coherence, structure, grammar, and overall effectiveness. Ensure that your suggestions are detailed and aimed at improving the paragraph within the context of the entire Document.";
            Word.Paragraphs paragraphs = CommonUtils.GetActiveDocument().Paragraphs;

            bool hasCommented = false; 
            if (Globals.ThisAddIn.Application.Selection.End - Globals.ThisAddIn.Application.Selection.Start > 0)
            {
                var selectionRange = CommonUtils.GetSelectionRange();
                try
                {
                    await CommentHandler.AddComment(CommonUtils.GetComments(), selectionRange, Review(paragraphs, selectionRange, prompt));
                } catch (OperationCanceledException ex)
                {
                    CommonUtils.DisplayWarning(ex);
                }
                hasCommented = true;
            }
            else
            {
                Word.Document document = CommonUtils.GetActiveDocument(); // Hash code of the active document gets changed after each comment!
                foreach (Word.Paragraph p in paragraphs)
                    // It isn't a paragraph if it doesn't contain a full stop.
                    if (p.Range.Text.Contains('.'))
                    {
                        await CommentHandler.AddComment(CommonUtils.GetComments(), p.Range, Review(paragraphs, p.Range, prompt, document));
                        hasCommented = true;
                    }
            }
            if (!hasCommented)
                MessageBox.Show(CultureHelper.GetLocalizedString("[ReviewButton_Click] MessageBox #1 (text)"), CultureHelper.GetLocalizedString("[ReviewButton_Click] MessageBox #1 (caption)"), MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            var range = (selectionRange.End - selectionRange.Start > 0) ? selectionRange : throw new InvalidRangeException(CultureHelper.GetLocalizedString("[AnalyzeText] InvalidRangeException #1"));
            
            ChatClient client = new ChatClient(ThisAddIn.Model, new ApiKeyCredential(ThisAddIn.ApiKey), ThisAddIn.ClientOptions);
            var streamingAnswer = client.CompleteChatStreamingAsync(
                new List<ChatMessage>() { new SystemChatMessage(systemPrompt), new UserChatMessage(@$"{userPrompt}: {range.Text}") },
                new ChatCompletionOptions() { MaxOutputTokenCount = ThisAddIn.ContextLength },
                ThisAddIn.CancellationTokenSource.Token
            );

            range.Delete();
            try
            {
                await AddStreamingContentToRange(streamingAnswer, range);
            } catch (OperationCanceledException ex)
            {
                CommonUtils.DisplayWarning(ex);
            }
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

            range.Text = WordMarkdown.RemoveMarkdownSyntax(response.ToString());
            WordMarkdown.ApplyAllMarkdownFormatting(range, response.ToString());
        }

        public static void CancelButtonVisibility(bool option)
        {
            _optionsBox.Visible = option;
        }

        private void UpdateCheckbox()
        {
            DefaultCheckBox.Checked = (Properties.Settings.Default.DefaultModel == ThisAddIn.Model);
        }

        private static AsyncCollectionResult<StreamingChatCompletionUpdate> Review(Word.Paragraphs context, Word.Range p, string prompt, Word.Document doc = null)
        {
            var docRange = Globals.ThisAddIn.Application.ActiveDocument.Range();
            List<UserChatMessage> userChat = new List<UserChatMessage>()
            {
                new UserChatMessage($@"Please review the following paragraph extracted from the Document: ""{CommonUtils.SubstringTokens(p.Text, (int)(ThisAddIn.ContextLength * 0.2))}"""),
                new UserChatMessage(prompt)
            };
            return RAGControl.AskQuestion(CommentSystemPrompt, userChat, docRange, doc);
        }
    }
}
