using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using HyperVectorDB.Embedder;
using Microsoft.Office.Tools;
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
        public static CancellationTokenSource CancellationTokenSource { get { return _cancellationTokenSource; } set { _cancellationTokenSource = value; } }
        public static bool IsAddinInitialized { get { return _isAddinInitialized; } set { _isAddinInitialized = value; } }
        public static OpenAIModelCollection ModelList { get { return _modelList; } }
        public static Dictionary<Word.Document, Tuple<CustomTaskPane, CustomTaskPane, RAGControl>> AllTaskPanes {  get { return _allTaskPanes; } }


        // Private
        private static string _openAIEndpoint = "http://localhost:11434/v1"; // Ollama endpoint
        private static string _apiKey = "dummy_key";
        private static string _model = "gpt-4o";
        private static string _embedModel = string.Empty;
        private static int _contextLength = ModelProperties.BaselineContextWindowLength;
        private static OpenAIClientOptions _clientOptions;
        private static EmbedderOpenAI _embedder;
        private static CancellationTokenSource _cancellationTokenSource = new();
        private static bool _isAddinInitialized = false;
        private static OpenAIModelCollection _modelList;
        private static Dictionary<Word.Document, Tuple<CustomTaskPane, CustomTaskPane, RAGControl>> _allTaskPanes = new();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Thread startup = new Thread(InitializeAddInStartup);
                startup.SetApartmentState(ApartmentState.STA);
                startup.Start();
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
                ((Word.ApplicationEvents4_Event)this.Application).NewDocument -= new Word.ApplicationEvents4_NewDocumentEventHandler(Application_NewDocument);
                ((Word.ApplicationEvents4_Event)this.Application).DocumentOpen -= new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
                ((Word.ApplicationEvents4_Event)this.Application).DocumentChange -= new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
                ((Word.ApplicationEvents4_Event)this.Application).WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(CommentHandler.Document_CommentsEventHandler);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        // https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb264456(v=office.12)?redirectedfrom=MSDN
        private async void Application_NewDocument(Word.Document doc)
        {
            try
            {
                if (this.Application.ShowWindowsInTaskbar)
                    await AddTaskPanesAsync(doc);
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }
        
        private async void Application_DocumentOpen(Word.Document doc)
        {
            try
            {
                if (this.Application.ShowWindowsInTaskbar)
                    await AddTaskPanesAsync(doc);
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void Application_DocumentChange()
        {
            try
            {
                RemoveClosedWindowTaskPanes();
                RemoveClosedWindowObjects();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void RemoveClosedWindowTaskPanes()
        {
            for (int i = this.CustomTaskPanes.Count - 1; i >= 0; i--)
                if (this.CustomTaskPanes[i].Window == null)
                    this.CustomTaskPanes.RemoveAt(i);
        }

        private void RemoveClosedWindowObjects()
        {
            for (int i = _allTaskPanes.Count - 1; i >= 0; i--)  // Iterate backwards to allow safe removal
            {
                var entry = _allTaskPanes.ElementAt(i);
                bool disposed = false;

                try
                {
                    // Check if the document's windows are all closed
                    if (entry.Key.Windows.Count == 0)
                    {
                        entry.Value.Item3.Dispose(); // Dispose of RAGControl
                        disposed = true;             // Mark as disposed
                        _allTaskPanes.Remove(entry.Key); // Remove the entry
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Handle the case where the document is already deleted
                    if (!disposed)  // Dispose if it hasn't been disposed yet
                    {
                        entry.Value.Item3.Dispose();
                    }

                    // Remove the entry after handling the exception
                    _allTaskPanes.Remove(entry.Key);
                }
            }
        }

        public static void AddTaskPanes(Word.Document doc)
        {
            if (_allTaskPanes.ContainsKey(doc)) return;
            AddCustomTaskPanesToWindow(doc, new RAGControl());
        }

        public static async Task AddTaskPanesAsync(Word.Document doc)
        {
            if (_allTaskPanes.ContainsKey(doc)) return;

            RAGControl ragControl = new RAGControl();

            // Capture the correct TaskScheduler for the UI thread
            var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();

            // The UI is a little more snappier?
            await Task.Factory.StartNew(() =>
            {
                AddCustomTaskPanesToWindow(doc, ragControl);
            }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);  // Use the UI thread scheduler
        }

        private static void AddCustomTaskPanesToWindow(Word.Document doc, RAGControl ragControl)
        {
            _allTaskPanes[doc] = new Tuple<CustomTaskPane, CustomTaskPane, RAGControl>(
                    Globals.ThisAddIn.CustomTaskPanes.Add(new GenerateUserControl(), Forge.CultureHelper.GetLocalizedString("this.GenerateButton.Label"), doc.ActiveWindow),
                    Globals.ThisAddIn.CustomTaskPanes.Add(ragControl, Forge.CultureHelper.GetLocalizedString("this.RAGControlButton.Label"), doc.ActiveWindow),
                    ragControl
            );
        }


        private void InitializeAddInStartup()
        {
            try
            {
                ((Word.ApplicationEvents4_Event)this.Application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(Application_NewDocument);
                ((Word.ApplicationEvents4_Event)this.Application).DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
                ((Word.ApplicationEvents4_Event)this.Application).DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
                ((Word.ApplicationEvents4_Event)this.Application).WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(CommentHandler.Document_CommentsEventHandler);
                lock (Forge.InitializeDoor)
                {
                    if (!_isAddinInitialized)
                        InitializeAddIn();
                }
            } catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        public static void InitializeAddIn()
        {
            InitializeEnvironmentVariables();
            _embedder = new EmbedderOpenAI(_embedModel, _apiKey, _clientOptions);
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

                // If no model was found, throw an exceptiond
                if (string.IsNullOrEmpty(_embedModel))
                    throw new ArgumentException(Forge.CultureHelper.GetLocalizedString("(ThisAddIn.cs) [InitializeAddIn] ArgumentException #1"));
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