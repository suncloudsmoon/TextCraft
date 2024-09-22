using System;
using System.Collections.Generic;
using System.Linq;
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
            // new models
            { "minicpm-v", 32768 },
            { "reader-lm", 256000 },
            { "mistral-small", 131072 },
            { "bespoke-minicheck", 32768 },
            { "qwen2.5", 32768 },
            { "nemotron-mini", 4096 },
            { "solar-pro", 4096 },
            { "qwen2.5-coder", 32768 }
        };

        public static int GetContextLength(string modelName)
        {
            if (openAIModelsContextLength.ContainsKey(modelName))
            {
                return openAIModelsContextLength[modelName];
            }
            else if (modelName.Contains(':'))
            {
                string key = modelName.Split(':')[0];
                return ollamaModelsContextLength.ContainsKey(key) ? ollamaModelsContextLength[key] : BaselineContextWindowLength;
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

        public static IEnumerable<string> GetModelList(ModelClient client)
        {
            OpenAIModelInfoCollection availableModels = client.GetModels().Value;
            return availableModels.Select(info => info.Id).ToList();
        }
    }
}
