using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using HyperVectorDB.Embedder;
using Microsoft.Office.Interop.Word;
using OpenAI;
using OpenAI.Chat;
using OpenAI.Models;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class ThisAddIn
    {
        // Public
        public const int BaselineContextWindowLength = 4096;
        public static string OpenAIEndpoint { get { return _openAIEndpoint; } }
        public static string ApiKey { get { return _apiKey; } }
        public static string Model { get { return _model; } set { _model = value; } }
        public static string EmbedModel { get { return _embedModel; } }
        public static int ContextLength { get { return _contextLength; } set { _contextLength = value; } }
        public static OpenAIClientOptions ClientOptions { get { return _clientOptions; } }
        public static EmbedderOpenAI Embedder { get { return _embedder; } }
        public static RAGControl RagControl { get { return _ragControl; } }
        public static CancellationTokenSource CancellationTokenSource { get { return _cancellationTokenSource; } set { _cancellationTokenSource = value; } }
        public static List<string> EmbedModels { get { return _embedModels; } }
        public static bool IsAddinInitialized { get { return _isAddinInitialized; } set { _isAddinInitialized = value; } }
        public static List<string> ModelList { get { return _modelList; } }

        // Private
        private static string _openAIEndpoint = "http://localhost:11434/v1";
        private static string _apiKey = "dummy_key";
        private static string _model = "gpt-4o";
        private static string _embedModel = string.Empty;
        private static int _contextLength = BaselineContextWindowLength;
        private static OpenAIClientOptions _clientOptions;
        private static EmbedderOpenAI _embedder;
        private static CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        private static RAGControl _ragControl;
        private static bool _isAddinInitialized = false;
        private static List<string> _modelList;

        private int _prevNumComments = 0;
        private bool _isDraftingComment = false;

        private static readonly List<string> _embedModels = new List<string>()
        {
            "all-minilm",
            "bge-m3",
            "bge-large",
            "paraphrase-multilingual"
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Document_CommentsEventHandler);
                if (!_isAddinInitialized)
                    InitializeAddin();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Document_CommentsEventHandler);
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        public static void InitializeAddin()
        {
            InitializeEnvironmentVariables();
            _embedder = new EmbedderOpenAI(_embedModel, _apiKey, _clientOptions);
            _ragControl = new RAGControl();
            _isAddinInitialized = true;
        }

        private async void Document_CommentsEventHandler(Word.Selection selection)
        {
            try
            {
                // For preventing unnecessary iteration of comments every time something changes in Word.
                int numComments = this.Application.ActiveDocument.Comments.Count;
                if (numComments != _prevNumComments)
                {
                    Comments comments = this.Application.ActiveDocument.Comments;

                    List<Comment> topLevelAIComments = new List<Comment>();
                    foreach (Comment c in comments)
                        if (c.Ancestor == null && c.Author == _model)
                            topLevelAIComments.Add(c);

                    foreach (Comment c in topLevelAIComments)
                    {
                        if (c.Replies.Count == 0) continue;
                        if (c.Replies[c.Replies.Count].Author != _model && !_isDraftingComment)
                        {
                            _isDraftingComment = true;

                            List<ChatMessage> chatHistory = new List<ChatMessage>();
                            chatHistory.Add(new UserChatMessage($@"Please review the following paragraph extracted from the Document: ""{c.Range.Text}"""));
                            chatHistory.Add(new UserChatMessage($@"Based on the previous AI comments, suggest additional specific improvements to the paragraph, focusing on clarity, coherence, structure, grammar, and overall effectiveness. Ensure that your suggestions are detailed and aimed at improving the paragraph within the context of the entire Document."));
                            for (int i = 1; i <= c.Replies.Count; i++)
                            {
                                Comment reply = c.Replies[i];
                                chatHistory.Add((i % 2 == 1) ? new UserChatMessage(reply.Range.Text) : new AssistantChatMessage(reply.Range.Text));
                            }

                            await Forge.AddComment(
                                c.Replies,
                                c.Range,
                                RAGControl.AskQuestion(Forge.SystemPrompt, chatHistory, this.Application.ActiveDocument.Range())
                            );
                            numComments++;

                            _isDraftingComment = false;
                        }
                    }
                    _prevNumComments = numComments;
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static void InitializeEnvironmentVariables()
        {
            GetEnvironmentVariableIfAvailable(ref _openAIEndpoint, "TEXTFORGE_OPENAI_ENDPOINT");
            GetEnvironmentVariableIfAvailable(ref _apiKey, "TEXTFORGE_API_KEY");
            GetEnvironmentVariableIfAvailable(ref _embedModel, "TEXTFORGE_EMBED_MODEL");

            // Initialize client variables
            _clientOptions = new OpenAIClientOptions
            {
                Endpoint = new Uri(_openAIEndpoint),
                ProjectId = "Operation Clippy",
                ApplicationId = "TextForge"
            };
            ModelClient modelRetriever = new ModelClient(_apiKey, _clientOptions);
            _modelList = GetModels(modelRetriever); // caching the response

            string defaultModel = Properties.Settings.Default.DefaultModel;
            _model = _modelList.Contains(defaultModel) ? defaultModel : modelRetriever.GetModels().Value[0].Id;
            _contextLength = RAGControl.GetContextLength(_model);

            // Set embed model
            if (string.IsNullOrEmpty(_embedModel))
            {
                foreach (var model in _modelList)
                {
                    if (model.Contains("embed"))
                    {
                        _embedModel = model;
                        break;
                    }
                    foreach (var item in _embedModels)
                        if (model.Contains(item))
                        {
                            _embedModel = model;
                            break;
                        }
                }
                if (string.IsNullOrEmpty(_embedModel))
                    throw new ArgumentException("Embed model is not installed on the computer!");
            }

        }

        private static void GetEnvironmentVariableIfAvailable(ref string dest, string variable)
        {
            var key = Environment.GetEnvironmentVariable(variable);
            if (key != null)
                dest = key;
        }
        private static List<string> GetModels(ModelClient model)
        {
            List<string> models = new List<string>();
            foreach (OpenAIModelInfo info in model.GetModels().Value)
                models.Add(info.Id);
            return models;
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}