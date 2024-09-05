using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Exceptions;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class RAGControl : UserControl
    {
        public HyperVectorDB.HyperVectorDB DB { get { return _db; } }
        public bool IsIndexing {  get { return _isIndexing; } }

        private BindingList<string> _fileList;
        private Queue<string> _removalQueue = new Queue<string>();
        private HyperVectorDB.HyperVectorDB _db;
        private bool _isIndexing;

        public const int CHUNK_LEN = 256;

        public RAGControl()
        {
            InitializeComponent();
            _fileList = new BindingList<string>();
            FileListBox.DataSource = _fileList;

            _db = new HyperVectorDB.HyperVectorDB(ThisAddIn.Embedder, Path.GetTempPath());
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
                    foreach (var filePath in filesToIndex)
                    {
                        await IndexDocumentAsync(filePath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AutoHideRemoveButton()
        {
            if (_fileList.Count == 0)
                RemoveButton.Enabled = false;
        }

        public string GetRAGContext(string query, int maxLen)
        {
            if (_fileList.Count == 0) return string.Empty;
            var result = _db.QueryCosineSimilarity(query, _fileList.Count);
            StringBuilder ragContext = new StringBuilder();
            foreach (var document in result.Documents)
                ragContext.AppendLine(document.DocumentString);
            return ragContext.ToString().Substring(0, (ragContext.Length >= maxLen) ? maxLen : ragContext.Length);
        }

        private async Task IndexDocumentAsync(string filePath)
        {
            List<string> fileContent;
            try
            {
                fileContent = await ReadPdfFileAsync(filePath, CHUNK_LEN);
            } finally
            {
                this.Invoke((MethodInvoker)delegate
                {
                    _fileList.Remove(filePath);
                    AutoHideRemoveButton();
                });
            }
            _db.CreateIndex(filePath);

            _isIndexing = true;
            await Task.Run(() => {
                _isIndexing = true;
                foreach (string chunk in fileContent)
                    _db.IndexDocument(chunk, null, null, filePath);
                _isIndexing = false;

                // Process any queued removal requests
                ProcessRemovalQueue();
            });
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

        public static async Task<List<string>> ReadPdfFileAsync(string filePath, int chunkLen)
        {
            return await Task.Run(() =>
            {
                List<string> chunks = new List<string>();
                try
                {
                    using (PdfDocument document = PdfDocument.Open(filePath))
                    {
                        foreach (var page in document.GetPages())
                        {
                            chunks.AddRange(SplitString(page.Text, chunkLen));
                        }
                    }
                } catch (PdfDocumentEncryptedException)
                {
                    // Testing
                    PasswordPrompt p = new PasswordPrompt();
                    if (p.ShowDialog() == DialogResult.OK)
                    {
                        using (PdfDocument document = PdfDocument.Open(filePath, new ParsingOptions { Password = p.Password }))
                        {
                            foreach (var page in document.GetPages())
                            {
                                chunks.AddRange(SplitString(page.Text, chunkLen));
                            }
                        }
                    } else
                    {
                        throw new ArgumentException("Invalid password dialog input!");
                    }
                }
                return chunks;
            });
        }

        // UTILS
        public static List<string> SplitString(string str, int chunkSize)
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

        public static string SubstringWithoutBounds(string text, int maxLen)
        {
            return (maxLen >= text.Length) ? text : text.Substring(0, maxLen);
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
    }
}
