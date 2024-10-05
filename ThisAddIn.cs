using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using HyperVectorDB.Embedder;
using OpenAI;
using OpenAI.Models;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class ThisAddIn
    {
        // Public
        public static string OpenAIEndpoint { get { return _openAIEndpoint; } }
        public static string ApiKey { get { return _apiKey; } }
        public static string Model { get { return _model; } set { _model = value; } }
        public static string EmbedModel { get { return _embedModel; } }
        public static int ContextLength { get { return _contextLength; } set { _contextLength = value; } }
        public static OpenAIClientOptions ClientOptions { get { return _clientOptions; } }
        public static EmbedderOpenAI Embedder { get { return _embedder; } }
        public static RAGControl RagControl { get { return _ragControl; } }
        public static CancellationTokenSource CancellationTokenSource { get { return _cancellationTokenSource; } set { _cancellationTokenSource = value; } }
        public static bool IsAddinInitialized { get { return _isAddinInitialized; } set { _isAddinInitialized = value; } }
        public static OpenAIModelCollection ModelList { get { return _modelList; } }

        // Private
        private static string _openAIEndpoint = "http://localhost:11434/v1"; // Ollama endpoint
        private static string _apiKey = "dummy_key";
        private static string _model = "gpt-4o";
        private static string _embedModel = string.Empty;
        private static int _contextLength = ModelProperties.BaselineContextWindowLength;
        private static OpenAIClientOptions _clientOptions;
        private static EmbedderOpenAI _embedder;
        private static CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();
        private static RAGControl _ragControl;
        private static bool _isAddinInitialized = false;
        private static OpenAIModelCollection _modelList;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(CommentHandler.Document_CommentsEventHandler);
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
                Properties.Settings.Default.Save();
                this.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(CommentHandler.Document_CommentsEventHandler);
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

        private static void InitializeEnvironmentVariables()
        {
            CommonUtils.GetEnvironmentVariableIfAvailable(ref _openAIEndpoint, "TEXTCRAFT_OPENAI_ENDPOINT");
            CommonUtils.GetEnvironmentVariableIfAvailable(ref _apiKey, "TEXTCRAFT_API_KEY");
            CommonUtils.GetEnvironmentVariableIfAvailable(ref _embedModel, "TEXTCRAFT_EMBED_MODEL");

            // Initialize client variables
            _clientOptions = new OpenAIClientOptions
            {
                Endpoint = new Uri(_openAIEndpoint),
                ProjectId = "Operation Clippy",
                UserAgentApplicationId = "TextCraft"
            };
            OpenAIModelClient modelRetriever = new OpenAIModelClient(new ApiKeyCredential(_apiKey), _clientOptions);
            _modelList = modelRetriever.GetModels().Value;

            string defaultModel = Properties.Settings.Default.DefaultModel;
            _model = _modelList.Any(model => model.Id == defaultModel) ? defaultModel : _modelList.First().Id;
            _contextLength = ModelProperties.GetContextLength(_model, _modelList);

            // Set embed model
            SetEmbedModelAutomatically();
        }

        private static void SetEmbedModelAutomatically()
        {
            if (string.IsNullOrEmpty(_embedModel))
            {
                // Use LINQ to find the first model that meets the condition
                _embedModel = _modelList.FirstOrDefault(model =>
                    model.Id.Contains("embed") || ModelProperties.UniqueEmbedModels.Any(item => model.Id.Contains(item))
                )?.Id;

                // If no model was found, throw an exception
                if (string.IsNullOrEmpty(_embedModel))
                    throw new ArgumentException("Embed model is not installed on the computer!");
            }
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