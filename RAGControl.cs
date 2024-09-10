using System;
using System.ClientModel;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenAI.Chat;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class RAGControl : UserControl
    {
        // Public
        public HyperVectorDB.HyperVectorDB DB { get { return _db; } }
        public bool IsIndexing {  get { return _isIndexing; } }
        public static readonly int CHUNK_LEN = TokensToCharCount(256);

        // Private
        private Queue<string> _removalQueue = new Queue<string>();
        private ConcurrentDictionary<int, int> _indexFileCount = new ConcurrentDictionary<int, int>();
        private BindingList<string> _fileList;
        private HyperVectorDB.HyperVectorDB _db;
        private bool _isIndexing;
        private readonly object progressBarLock = new object();

        private static readonly Dictionary<string, int> modelsContextLength = new Dictionary<string, int>()
        {
            { "gpt-4-0125-preview", 128000 },
            { "gpt-4-1106-preview", 128000 },
            { "gpt-3.5-turbo-instruct", 4096 },
        };

        private static readonly Dictionary<string, int> ollamaModelsContextLength = new Dictionary<string, int>()
        {
            { "llama3.1", 131072 },
            { "gemma2", 8192 },
            { "mistral-nemo", 1024000 },
            { "mistral-large", 32768 },
            { "qwen2", 32768 },
            { "deepseek-coder-v2", 163840 },
            { "phi3", 4096 }, // the lowest config of phi3 (https://ollama.com/library/phi3:medium-4k/blobs/0d98d611d31b) 
            { "mistral", 32768 },
            { "mixtral", 32768 },
            { "codegemma", 8192 },
            { "command-r", 131072 },
            { "command-r-plus", 131072 },
            { "llava", 4096 }, // https://ollama.com/library/llava:13b-v1.6-vicuna-q4_0/blobs/87d5b13e5157
            { "llama3", 8192 },
            { "gemma", 8192 },
            { "qwen", 32768 },
            { "llama2", 4096 },
            { "codellama", 16384 },
            { "dolphin-mixtral", 32768 },
            { "phi", 2048 },
            { "llama2-uncensored", 2048 },
            { "deepseek-coder", 16384 },
            { "dolphin-mistral", 32768 },
            { "zephyr", 32768 },
            { "starcoder2", 16384 },
            { "dolphin-llama3", 8192 },
            { "orca-mini", 2048 },
            { "yi", 4096 },
            { "mistral-openorca", 32768 },
            { "llava-llama3", 8192 },
            { "starcoder", 8192 },
            { "llama2-chinese", 4096 },
            { "vicuna", 2048 }, // https://ollama.com/library/vicuna:7b-q3_K_S/blobs/a887a7755a1e
            { "tinyllama", 2048 },
            { "codestral", 32768 },
            { "wizard-vicuna-uncensored", 2048 },
            { "nous-hermes2", 4096 },
            { "openchat", 8192 },
            { "aya", 8192 },
            { "granite-code", 2048 },
            { "wizardlm2", 32768 },
            { "tinydolphin", 4096 },
            { "wizardcoder", 16384 },
            { "stable-code", 16384 },
            { "openhermes", 32768 },
            { "codeqwen", 65536 },
            { "codegeex4", 131072 },
            { "stablelm2", 4096 },
            { "wizard-math", 2048 }, // https://ollama.com/library/wizard-math:7b-q5_1/blobs/b39818bfe610
            { "qwen2-math", 4096 },
            { "neural-chat", 32768 },
            { "llama3-gradient", 1048576 },
            { "phind-codellama", 16384 },
            { "nous-hermes", 2048 }, // https://ollama.com/library/nous-hermes:13b-q4_1/blobs/433b4abded9b
            { "sqlcoder", 8192 }, // https://ollama.com/library/sqlcoder:15b-q8_0/blobs/2319eec9425f
            { "dolphincoder", 16384 },
            { "xwinlm", 4096 },
            { "deepseek-llm", 4096 },
            { "yarn-llama2", 65536 },
            { "llama3-chatqa", 8192 },
            { "wizardlm", 2048 },
            { "starling-lm", 8192 },
            { "falcon", 2048 },
            { "orca2", 4096 },
            { "moondream", 2048 },
            { "samantha-mistral", 32768 },
            { "solar", 4096 },
            { "smollm", 2048 },
            { "stable-beluga", 4096 },
            { "dolphin-phi", 2048 },
            { "deepseek-v2", 163840 },
            { "glm4", 8192 }, // https://ollama.com/library/glm4:9b-text-q6_K/blobs/2f657a57a8df
            { "phi3.5", 131072 },
            { "bakllava", 32768 },
            { "wizardlm-uncensored", 4096 },
            { "yarn-mistral", 32768 }, // https://ollama.com/library/yarn-mistral/blobs/0e8703041ff2
            { "medllama2", 4096 },
            { "llama-pro", 4096 },
            { "llava-phi3", 4096 },
            { "meditron", 2048 },
            { "nous-hermes2-mixtral", 32768 },
            { "nexusraven", 16384 },
            { "hermes3", 131072 },
            { "codeup", 4096 },
            { "everythinglm", 16384 },
            { "internlm2", 32768 },
            { "magicoder", 16384 },
            { "stablelm-zephyr", 4096 },
            { "codebooga", 16384 },
            { "yi-coder", 131072 },
            { "mistrallite", 32768 },
            { "llama3-groq-tool-use", 8192 },
            { "falcon2", 2048 },
            { "wizard-vicuna", 2048 },
            { "duckdb-nsql", 16384 },
            { "megadolphin", 4096 },
            { "reflection", 8192 },
            { "notux", 32768 },
            { "goliath", 4096 },
            { "open-orca-platypus2", 4096 },
            { "notus", 32768 },
            { "dbrx", 32768 },
            { "mathstral", 32768 },
            { "alfred", 2048 },
            { "nuextract", 4096 },
            { "firefunction-v2", 8192 },
            { "deepseek-v2.5", 163840 },
        };

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
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                    openFileDialog.Multiselect = true;
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
                    }

                    ChangeProgressBarVisibility(true);
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

                    await ChangeProgressBarAfterSleep(10, false);
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

            _isIndexing = true;
            await Task.Run(() => {
                _isIndexing = true;
                int count = fileContent.Count();
                for (int i = 0; i < count; i++) {
                    _db.IndexDocument(fileContent.ElementAt(i), null, null, filePath);
                    this.Invoke((MethodInvoker)delegate
                    {
                        UpdateProgressBar( (i + 1) / count );
                    });
                }
                _isIndexing = false;

                // Process any queued removal requests
                ProcessRemovalQueue();
            });
        }

        private async Task ChangeProgressBarAfterSleep(int seconds, bool val)
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
            return this.progressBar1.Value / 100f;
        }

        private void SetProgressBarValue(float val)
        {
            this.progressBar1.Value = (int)(val * 100);
        }

        private void UpdateProgressBar(float val)
        {
            lock (progressBarLock)
            {
                int fileCount = 0;
                foreach (var count in _indexFileCount)
                    fileCount += count.Value;
                this.progressBar1.Value += (int)((val * 100) / fileCount);
            }
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

        private bool RemoveDocument(string filePath)
        {
            return _db.DeleteIndex(filePath);
        }

        public static async Task<IEnumerable<string>> ReadPdfFileAsync(string filePath, int chunkLen)
        {
            return await Task.Run(() =>
            {
                List<string> chunks = new List<string>();
                PdfDocument doc = PdfDocument.Open(filePath);
                if (doc.IsEncrypted)
                {
                    PasswordPrompt passwordDialog = new PasswordPrompt();
                    if (passwordDialog.ShowDialog() == DialogResult.OK)
                    {
                        PdfDocument unlockedDoc = PdfDocument.Open(filePath, new ParsingOptions { Password = passwordDialog.Password });
                        IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen);
                        unlockedDoc.Dispose();
                    }
                    else
                    {
                        throw new ArgumentException("Could not unlock PDF file due to incorrect password!");
                    }
                }
                else
                {
                    IteratePdfFile(ref doc, ref chunks, chunkLen);
                }
                doc.Dispose();

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
                        PdfDocument embedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray());
                        if (embedDoc.IsEncrypted)
                        {
                            PasswordPrompt passwordDialog = new PasswordPrompt();
                            if (passwordDialog.ShowDialog() == DialogResult.OK)
                            {
                                PdfDocument unlockedDoc = PdfDocument.Open(embeddedFile.Bytes.ToArray(), new ParsingOptions { Password = passwordDialog.Password });
                                IteratePdfFile(ref unlockedDoc, ref chunks, chunkLen);
                                unlockedDoc.Dispose();
                            }
                            else
                            {
                                throw new ArgumentException("Could not unlock PDF file due to incorrect password!");
                            }
                        }
                        else
                        {
                            IteratePdfFile(ref embedDoc, ref chunks, chunkLen);
                        }
                        embedDoc.Dispose();
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
                    chunks.AddRange(SplitString(block.Text, chunkLen));
            }
        }

        public string GetRAGContext(string query, int maxTokens)
        {
            if (_fileList.Count == 0) return string.Empty;
            var result = _db.QueryCosineSimilarity(query, _fileList.Count * 5); // 5 results per file
            StringBuilder ragContext = new StringBuilder();
            foreach (var document in result.Documents)
                ragContext.AppendLine(document.DocumentString);
            return SubstringTokens(ragContext.ToString(), maxTokens);
        }

        // UTILS
        public static AsyncCollectionResult<StreamingChatCompletionUpdate> AskQuestion(SystemChatMessage systemPrompt, IEnumerable<ChatMessage> messages, Word.Range context)
        {
            string document = context.Text;

            int userPromptLen = 0;
            foreach (var m in messages)
                userPromptLen += m.Content[0].Text.Length;

            ChatMessage lastUserPrompt = messages.Last();

            var constraints = RAGControl.OptimizeConstraint(
                0.8f,
                ThisAddIn.ContextLength,
                RAGControl.CharToTokenCount(systemPrompt.Content[0].Text.Length + userPromptLen),
                RAGControl.CharToTokenCount(document.Length)
            );
            if (constraints["document_content_rag"] == 1f)
                document = RAGControl.GetWordDocumentAsRAG(lastUserPrompt.Content[0].Text, context);

            string ragQuery =
                (constraints["rag_context"] == 0f) ? string.Empty : ThisAddIn.RagControl.GetRAGContext(lastUserPrompt.Content[0].Text, (int)(ThisAddIn.ContextLength * constraints["rag_context"]));

            List<ChatMessage> chat = new List<ChatMessage>();
            chat.Add(systemPrompt);
            chat.Add(new UserChatMessage(
                $@"Document Content: ""{RAGControl.SubstringTokens(document, (int)(ThisAddIn.ContextLength * constraints["document_content"]))}"""
            ));
            if (ragQuery != string.Empty)
                chat.Add(new UserChatMessage($@"RAG Context: ""{ragQuery}"""));
            chat.AddRange(messages);

            ChatClient client = new ChatClient(ThisAddIn.Model, ThisAddIn.ApiKey, ThisAddIn.ClientOptions);
            // https://github.com/ollama/ollama/pull/6504
            return client.CompleteChatStreamingAsync(
                chat,
                new ChatCompletionOptions() { MaxTokens = ThisAddIn.ContextLength },
                ThisAddIn.CancellationTokenSource.Token
            );
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
            var chunks = SplitString(context.Text, CHUNK_LEN);
            foreach (var chunk in chunks)
            {
                DB.IndexDocument(chunk);
            }
            var result = DB.QueryCosineSimilarity(query, 3);
            StringBuilder ragContext = new StringBuilder();
            foreach (var doc in result.Documents)
                ragContext.AppendLine(doc.DocumentString);
            return ragContext.ToString();
        }

        public static int GetContextLength(string model)
        {
            if (modelsContextLength.ContainsKey(model))
                return modelsContextLength[model];
            if (model.Contains(':'))
            {
                string key = model.Split(':')[0];
                if (ollamaModelsContextLength.ContainsKey(key))
                    return ollamaModelsContextLength[key];
                else
                    return ThisAddIn.BaselineContextWindowLength;
            }
            else if (modelsContextLength.ContainsKey(model))
            {
                return modelsContextLength[model];
            }
            else if (model.StartsWith("gpt-4-turbo"))
            {
                return 128000;
            }
            else if (model.StartsWith("gpt-4-mini"))
            {
                return 128000;
            }
            else if (model.StartsWith("gpt-4"))
            {
                return 8192;
            }
            else if (model.StartsWith("gpt-3.5-turbo"))
            {
                return 16385;
            }
            else
            {
                return ThisAddIn.BaselineContextWindowLength;
            }
        }

        public static IEnumerable<string> SplitString(string str, int chunkSize)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < str.Length; i += chunkSize)
            {
                if (i + chunkSize > str.Length)
                    chunkSize = str.Length - i;

                result.Add(str.Substring(i, chunkSize));
            }
            return result;
        }

        public static string SubstringTokens(string text, int maxTokens)
        {
            return SubstringWithoutBounds(text, TokensToCharCount(maxTokens));
        }

        private static string SubstringWithoutBounds(string text, int maxLen)
        {
            return (maxLen >= text.Length) ? text : text.Substring(0, maxLen);
        }

        public static int TokensToCharCount(int tokenCount)
        {
            return tokenCount * 4; // https://platform.openai.com/tokenizer
        }
        public static int CharToTokenCount(int tokenCount)
        {
            return tokenCount / 4; // https://platform.openai.com/tokenizer
        }
    }
}
