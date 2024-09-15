using System;
using System.ClientModel;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.Exceptions;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class RAGControl : UserControl
    {
        // Public
        public static readonly int CHUNK_LEN = CommonUtils.TokensToCharCount(256);

        // Private
        private Queue<string> _removalQueue = new Queue<string>();
        private ConcurrentDictionary<int, int> _indexFileCount = new ConcurrentDictionary<int, int>();
        private BindingList<string> _fileList;
        private HyperVectorDB.HyperVectorDB _db;
        private bool _isIndexing;
        private readonly object progressBarLock = new object();

        public RAGControl()
        {
            try
            {
                InitializeComponent();
                _fileList = new BindingList<string>();
                FileListBox.DataSource = _fileList;

                _db = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async void AddButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog() { Title = "Select Files for RAG", Filter = "PDF files (*.pdf)|*.pdf", Multiselect = true })
                {
                    List<string> filesToIndex = new List<string>();
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        foreach (string fileName in openFileDialog.FileNames)
                        {
                            if (!_fileList.Contains(fileName))
                            {
                                _fileList.Add(fileName);
                                filesToIndex.Add(fileName);
                                if (!RemoveButton.Enabled)
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        RemoveButton.Enabled = true;
                                    });
                                }
                            }
                        }

                        ChangeProgressBarVisibility(true);
                        {
                            int dictCount = _indexFileCount.Count;
                            _indexFileCount.TryAdd(dictCount, filesToIndex.Count);
                            lock (progressBarLock)
                            {
                                SetProgressBarValue(GetProgressBarValue() / (dictCount + 1));
                            }

                            foreach (var filePath in filesToIndex)
                            {
                                await IndexDocumentAsync(filePath);
                            }

                            int temp;
                            _indexFileCount.TryRemove(dictCount, out temp);
                        }
                        await ChangeProgressBarVisibilityAfterSleep(5, false);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void RemoveButton_Click(object sender, EventArgs e)
        {
            try
            {
                string selectedDocument = FileListBox.SelectedItem.ToString();
                if (_isIndexing)
                {
                    if (!_removalQueue.Contains(selectedDocument))
                        _removalQueue.Enqueue(selectedDocument);
                }
                else
                {
                    RemoveDocument(selectedDocument);
                }
                _fileList.RemoveAt(FileListBox.SelectedIndex);
                AutoHideRemoveButton();
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void AutoHideRemoveButton()
        {
            if (_fileList.Count == 0)
                RemoveButton.Enabled = false;
        }

        private async Task IndexDocumentAsync(string filePath)
        {
            IEnumerable<string> fileContent;
            try
            {
                fileContent = await ReadPdfFileAsync(filePath, CHUNK_LEN);
            } catch
            {
                this.Invoke((MethodInvoker)delegate
                {
                    _fileList.Remove(filePath);
                    AutoHideRemoveButton();
                });
                throw;
            }

            _db.CreateIndex(filePath);
            await Task.Run(() => {
                _isIndexing = true;

                foreach (var content in fileContent)
                    AddDocument(filePath, content);
                
                this.Invoke((MethodInvoker)delegate
                {
                    UpdateProgressBar(1);
                });
                _isIndexing = false;

                // Process any queued removal requests
                ProcessRemovalQueue();
            });
        }

        private async Task ChangeProgressBarVisibilityAfterSleep(int seconds, bool val)
        {
            await Task.Delay(seconds * 1000);
            this.Invoke((MethodInvoker)delegate
            {
                ChangeProgressBarVisibility(val);
                ResetProgressBar();
            });
        }

        private void ChangeProgressBarVisibility(bool val)
        {
            this.progressBar1.Visible = val;
        }

        private void ResetProgressBar()
        {
            SetProgressBarValue(0);
        }

        private float GetProgressBarValue()
        {
            return this.progressBar1.Value / ((float) this.progressBar1.Maximum);
        }

        private void SetProgressBarValue(float val)
        {
            this.progressBar1.Value = (int)(val * this.progressBar1.Maximum);
        }

        private void UpdateProgressBar(float val)
        {
            lock (progressBarLock)
            {
                int fileCount = GetIndexFileCount();

                // Clipping
                int maxProgress = this.progressBar1.Maximum;
                int incrementVal = (int)((val * maxProgress) / fileCount);
                if (incrementVal + this.progressBar1.Value > maxProgress)
                    this.progressBar1.Value = maxProgress;
                else
                    this.progressBar1.Value += incrementVal;
            }
        }

        private int GetIndexFileCount()
        {
            int fileCount = 0;
            foreach (var count in _indexFileCount)
                fileCount += count.Value;
            return fileCount;
        }

        private void ProcessRemovalQueue()
        {
            int initialQueueCount = _removalQueue.Count;
            for (int i = 0; i < initialQueueCount; i++)
            {
                string documentToRemove = _removalQueue.Dequeue();
                if (!_fileList.Contains(documentToRemove))
                {
                    if (!RemoveDocument(documentToRemove))
                    {
                        _removalQueue.Enqueue(documentToRemove);
                        break;
                    }
                }
            }
        }

        private bool AddDocument(string filePath, string content)
        {
            return _db.IndexDocument(filePath, content);
        }

        private bool RemoveDocument(string filePath)
        {
            return _db.DeleteIndex(filePath);
        }

        public static async Task<IEnumerable<string>> ReadPdfFileAsync(string filePath, int chunkLen)
        {
            return await Task.Run(() =>
            {
                List<string> chunks = new List<string>();
                try
                {
                    PdfDocument doc;
                    try { doc = PdfDocument.Open(filePath); }
                    catch (PdfDocumentEncryptedException) { throw new ArgumentException(); }

                    try { IteratePdfFile(ref doc, ref chunks, chunkLen); }
                    finally { doc.Dispose(); }
                }
                catch (ArgumentException)
                {
                    PasswordPrompt passwordDialog = new PasswordPrompt();
                    if (passwordDialog.ShowDialog() == DialogResult.OK)
                    {
                        PdfDocument unlockedDoc = PdfDocument.Open(filePath, new ParsingOptions { Password = passwordDialog.Password });
                        try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen); } 
                        finally { unlockedDoc.Dispose(); }
                    }
                    else
                    {
                        throw new InvalidDataException("Could not unlock PDF file due to incorrect password!");
                    }
                }
                return chunks;
            });
        }

        private static void IteratePdfFile(ref PdfDocument document, ref List<string> chunks, int chunkLen)
        {
            IterateInnerPdfFile(ref document, ref chunks, chunkLen);

            IReadOnlyList<EmbeddedFile> embeddedFiles;
            if (document.Advanced.TryGetEmbeddedFiles(out embeddedFiles))
            {
                foreach (var embeddedFile in embeddedFiles)
                {
                    if (embeddedFile.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            PdfDocument embedDoc;
                            try { embedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray()); }
                            catch (PdfDocumentEncryptedException) { throw new ArgumentException(); }

                            try { IteratePdfFile(ref embedDoc, ref chunks, chunkLen); }
                            finally { embedDoc.Dispose(); }
                        } catch (ArgumentException)
                        {
                            PasswordPrompt passwordDialog = new PasswordPrompt();
                            if (passwordDialog.ShowDialog() == DialogResult.OK)
                            {
                                PdfDocument unlockedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray(), new ParsingOptions { Password = passwordDialog.Password });
                                try { IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen); }
                                finally { unlockedDoc.Dispose(); }
                            }
                            else
                            {
                                throw new InvalidDataException("Could not unlock PDF file due to incorrect password!");
                            }
                        }
                    }
                }
            }
        }

        private static void IterateInnerPdfFile(ref PdfDocument doc, ref List<string> chunks, int chunkLen)
        {
            foreach (var page in doc.GetPages())
            {
                var blocks = DocstrumBoundingBoxes.Instance.GetBlocks(page.GetWords());
                foreach (var block in blocks)
                    chunks.AddRange(CommonUtils.SplitString(block.Text, chunkLen));
            }
        }

        public string GetRAGContext(string query, int maxTokens)
        {
            if (_fileList.Count == 0) return string.Empty;
            var result = _db.QueryCosineSimilarity(query, _fileList.Count * 5); // 5 results per file
            StringBuilder ragContext = new StringBuilder();
            foreach (var document in result.Documents)
                ragContext.AppendLine(document.DocumentString);
            return CommonUtils.SubstringTokens(ragContext.ToString(), maxTokens);
        }

        // UTILS
        public static AsyncCollectionResult<StreamingChatCompletionUpdate> AskQuestion(SystemChatMessage systemPrompt, IEnumerable<ChatMessage> messages, Word.Range context)
        {
            string document = context.Text;
            int userPromptLen = GetUserPromptLen(messages);
            ChatMessage lastUserPrompt = messages.Last();

            var constraints = RAGControl.OptimizeConstraint(
                0.8f,
                ThisAddIn.ContextLength,
                CommonUtils.CharToTokenCount(systemPrompt.Content[0].Text.Length + userPromptLen),
                CommonUtils.CharToTokenCount(document.Length)
            );
            if (constraints["document_content_rag"] == 1f)
                document = RAGControl.GetWordDocumentAsRAG(lastUserPrompt.Content[0].Text, context);

            string ragQuery =
                (constraints["rag_context"] == 0f) ? string.Empty : ThisAddIn.RagControl.GetRAGContext(lastUserPrompt.Content[0].Text, (int)(ThisAddIn.ContextLength * constraints["rag_context"]));

            List<ChatMessage> chatHistory = new List<ChatMessage>()
            {
                systemPrompt,
                new UserChatMessage($@"Document Content: ""{CommonUtils.SubstringTokens(document, (int)(ThisAddIn.ContextLength * constraints["document_content"]))}""")
            };
            if (ragQuery != string.Empty)
                chatHistory.Add(new UserChatMessage($@"RAG Context: ""{ragQuery}"""));
            chatHistory.AddRange(messages);

            ChatClient client = new ChatClient(ThisAddIn.Model, ThisAddIn.ApiKey, ThisAddIn.ClientOptions);
            // https://github.com/ollama/ollama/pull/6504
            return client.CompleteChatStreamingAsync(
                chatHistory,
                new ChatCompletionOptions() { MaxTokens = ThisAddIn.ContextLength },
                ThisAddIn.CancellationTokenSource.Token
            );
        }

        private static int GetUserPromptLen(IEnumerable<ChatMessage> messageList)
        {
            int userPromptLen = 0;
            foreach (var message in messageList)
                userPromptLen += message.Content[0].Text.Length;
            return userPromptLen;
        }

        public static Dictionary<string, float> OptimizeConstraint(float maxPercentage, int contextLength, int promptTokenLen, int documentContentTokenLen)
        {
            Dictionary<string, float> constraintPairs = new Dictionary<string, float>();
            if (promptTokenLen >= contextLength * 0.9)
            {
                constraintPairs["rag_context"] = 0f;
                constraintPairs["document_content"] = (float)(maxPercentage * 0.1);
                constraintPairs["document_content_rag"] = (documentContentTokenLen > contextLength * maxPercentage * 0.1) ? 1f : 0f;
            }
            else
            {
                float promptPercentage = (float)promptTokenLen / (float)contextLength;
                constraintPairs["rag_context"] = (float)( (1 - promptPercentage) * maxPercentage * 0.3);
                constraintPairs["document_content"] = (float)((1 - promptPercentage) * maxPercentage * 0.7);
                constraintPairs["document_content_rag"] = (documentContentTokenLen > contextLength * (1 - promptPercentage) * maxPercentage * 0.7) ? 1f : 0f;
            }
            return constraintPairs;
        }

        public static string GetWordDocumentAsRAG(string query, Word.Range context)
        {
            // Get RAG context
            HyperVectorDB.HyperVectorDB DB = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
            var chunks = CommonUtils.SplitString(context.Text, CHUNK_LEN);
            foreach (var chunk in chunks)
                DB.IndexDocument(chunk);

            var result = DB.QueryCosineSimilarity(query, CommonUtils.GetWordPageCount() * 3);
            StringBuilder ragContext = new StringBuilder();
            foreach (var doc in result.Documents)
                ragContext.AppendLine(doc.DocumentString);
            return ragContext.ToString();
        }
    }
}
