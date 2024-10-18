using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using OpenAI.Models;

namespace TextForge
{
    internal class ModelProperties
    {
        // Public
        public const int BaselineContextWindowLength = 4096; // Change this if necessary
        public static List<string> UniqueEmbedModels { get { return _embedModels; } }

        // Private
        private static readonly List<string> _embedModels = new List<string>()
        {
            "all-minilm",
            "bge-m3",
            "bge-large",
            "paraphrase-multilingual"
        };

        private static readonly Dictionary<string, int> openAIModelsContextLength = new Dictionary<string, int>()
        {
            { "gpt-4-0125-preview", 128000 },
            { "gpt-4-1106-preview", 128000 },
            { "gpt-3.5-turbo-instruct", 4096 },
        };

        private static bool IsOllamaEndpoint = false;
        private static bool IsOllamaFetched = false;
        private static Dictionary<string, int> ollamaContextWindowCache = new Dictionary<string, int>();
        private static readonly CultureLocalizationHelper _cultureHelper = new CultureLocalizationHelper("TextForge.Forge", typeof(Forge).Assembly);


        public static int GetContextLength(string modelName, OpenAIModelCollection availableModels)
        {
            if (openAIModelsContextLength.ContainsKey(modelName))
            {
                return openAIModelsContextLength[modelName];
            }
            else if (modelName.Contains(':'))
            {
                try
                {
                    if (!IsOllamaFetched)
                    {
                        IsOllamaEndpoint = IsOllama(availableModels);
                        IsOllamaFetched = true;
                    }
                    if (IsOllamaEndpoint)
                    {
                        int contextWindow;
                        if (!ollamaContextWindowCache.TryGetValue(modelName, out contextWindow))
                        {
                            contextWindow = GetOllamaModelContextWindow(modelName);
                            ollamaContextWindowCache[modelName] = contextWindow;
                        }
                        return contextWindow;
                    } else
                    {
                        return BaselineContextWindowLength;
                    }
                } catch (OllamaMissingContextWindowException ex)
                {
                    CommonUtils.DisplayWarning(ex);
                    return BaselineContextWindowLength;
                }
            }
            else if (modelName.StartsWith("o1"))
            {
                return 128000;
            }
            else if (modelName.StartsWith("gpt-4-turbo"))
            {
                return 128000;
            }
            else if (modelName.StartsWith("gpt-4-mini"))
            {
                return 128000;
            }
            else if (modelName.StartsWith("gpt-4"))
            {
                return 8192;
            }
            else if (modelName.StartsWith("gpt-3.5-turbo"))
            {
                return 16385;
            }
            else
            {
                return BaselineContextWindowLength;
            }
        }

        public static IEnumerable<string> GetModelList(OpenAIModelCollection availableModels)
        {
            return availableModels.Select(info => info.Id).ToList();
        }

        private static bool IsOllama(OpenAIModelCollection availableModels)
        {
            return (availableModels.Count == 0) ? false : availableModels.First().OwnedBy == "library";
        }

        private static int GetOllamaModelContextWindow(string model)
        {
            var ollamaEndpoint = ThisAddIn.OpenAIEndpoint.Replace("/v1", "");

            Ollama ollamaInstance = new Ollama(new Uri(ollamaEndpoint));
            var dict = ollamaInstance.Show(model, true).Result; // or await, if Show() is async

            // Navigate to "model_info"
            if (dict.TryGetValue("model_info", out var modelInfoObj) && modelInfoObj is JsonElement modelInfoElement)
            {
                // Use JsonNode or JsonElement to search for "context_length" key
                var modelInfoNode = JsonNode.Parse(modelInfoElement.GetRawText());

                foreach (var keyValuePair in modelInfoNode.AsObject())
                {
                    // Search for a nested object containing "context_length"
                    if (keyValuePair.Key.EndsWith(".context_length"))
                    {
                        return int.Parse(keyValuePair.Value.ToString());
                    }
                }
            }
            
            throw new OllamaMissingContextWindowException(string.Format(_cultureHelper.GetLocalizedString("(ModelProperties.cs) [GetContextLength] OllamaMissingContextWindowException #1"), model));
        }
    }

    public class OllamaMissingContextWindowException : ApplicationException
    {
        public OllamaMissingContextWindowException(string message) : base(message) { }
    }
}
