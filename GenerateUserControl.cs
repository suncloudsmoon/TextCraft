using System;
using System.ClientModel;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HyperVectorDB.Embedder;
using OpenAI.Chat;
using Word = Microsoft.Office.Interop.Word;

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
                var streamingAnswer = AskQuestion(_systemPrompt, Globals.ThisAddIn.Application.ActiveDocument.Range(), textBoxContent);

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

        public static AsyncCollectionResult<StreamingChatCompletionUpdate> AskQuestion(SystemChatMessage systemPrompt, Word.Range context, string prompt)
        {
            string document = (Globals.ThisAddIn.Application.ActiveDocument.Words.Count * 1.4 > ThisAddIn.ContextLength * 0.4) ? RAGControl.GetWordDocumentAsRAG(prompt, context) : context.Text;

            // 0.1 of context length leftover to account for UserChatMessage and other stuff
            SystemChatMessage systemPromptBounded = new SystemChatMessage(RAGControl.SubstringWithoutBounds(systemPrompt.Content[0].Text, (int)(ThisAddIn.ContextLength * 0.1)));
            UserChatMessage fullPrompt = new UserChatMessage($@"{RAGControl.SubstringWithoutBounds(prompt, (int) (ThisAddIn.ContextLength * 0.2))}{Environment.NewLine}RAG Context: ""{ThisAddIn.RagControl.GetRAGContext(prompt, (int)(ThisAddIn.ContextLength * 0.2))}""{Environment.NewLine}Document Content: ""{RAGControl.SubstringWithoutBounds(document, (int)(ThisAddIn.ContextLength * 0.4))}""");
            
            ChatClient client = new ChatClient(ThisAddIn.Model, ThisAddIn.ApiKey, ThisAddIn.ClientOptions);
            return client.CompleteChatStreamingAsync(new List<ChatMessage>() { systemPromptBounded, fullPrompt }, null, ThisAddIn.CancellationTokenSource.Token);
        }
    }
}
