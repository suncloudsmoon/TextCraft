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
        public static string OpenAIEndpoint { get { return _openAIEndpoint; } }
        public static string ApiKey { get { return _apiKey; } }
        public static string Model { get { return _model; } set { _model = value; } }
        public static string EmbedModel { get { return _embedModel; } }
        public static int ContextLength { get { return _contextLength; } }
        public static OpenAIClientOptions ClientOptions { get { return _clientOptions; } }
        public static EmbedderOpenAI Embedder { get { return _embedder; } }
        public static RAGControl RagControl { get { return _ragControl; } }
        public static CancellationTokenSource CancellationTokenSource { get { return _cancellationTokenSource; } set { _cancellationTokenSource = value; } }
        public static List<string> EmbedModels { get { return _embedModels; } }
        public static bool IsAddinInitialized { get { return _isAddinInitialized; } set { _isAddinInitialized = value; } }
        public static List<string> ModelList { get { return _modelList; } }

        // Private
        private static string _openAIEndpoint = "http://localhost:11434/";
        private static string _apiKey = "dummy_key";
        private static string _model = "gpt-4o";
        private static string _embedModel = string.Empty;
        private static int _contextLength = 4096; // Changed from 32768
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

        private static readonly Dictionary<string, string[]> Keywords = new Dictionary<string, string[]>
        {
            ["python"] = new[]
        {
            "False", "None", "True", "and", "as", "assert", "async", "await", "break", "class", "continue", "def",
            "del", "elif", "else", "except", "finally", "for", "from", "global", "if", "import", "in", "is", "lambda",
            "nonlocal", "not", "or", "pass", "raise", "return", "try", "while", "with", "yield"
        },
                ["c"] = new[]
        {
            "auto", "break", "case", "char", "const", "continue", "default", "do", "double", "else", "enum", "extern",
            "float", "for", "goto", "if", "int", "long", "register", "return", "short", "signed", "sizeof", "static",
            "struct", "switch", "typedef", "union", "unsigned", "void", "volatile", "while"
        },
                ["cpp"] = new[]
        {
            "alignas", "alignof", "and", "and_eq", "asm", "auto", "bitand", "bitor", "bool", "break", "case", "catch",
            "char", "char8_t", "char16_t", "char32_t", "class", "compl", "concept", "const", "consteval", "constexpr",
            "constinit", "const_cast", "continue", "co_await", "co_return", "co_yield", "decltype", "default", "delete",
            "do", "double", "dynamic_cast", "else", "enum", "explicit", "export", "extern", "false", "float", "for",
            "friend", "goto", "if", "inline", "int", "long", "mutable", "namespace", "new", "noexcept", "not", "not_eq",
            "nullptr", "operator", "or", "or_eq", "private", "protected", "public", "register", "reinterpret_cast",
            "requires", "return", "short", "signed", "sizeof", "static", "static_assert", "static_cast", "struct",
            "switch", "template", "this", "thread_local", "throw", "true", "try", "typedef", "typeid", "typename",
            "union", "unsigned", "using", "virtual", "void", "volatile", "wchar_t", "while", "xor", "xor_eq"
        },
                ["csharp"] = new[]
        {
            "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char", "checked", "class", "const",
            "continue", "decimal", "default", "delegate", "do", "double", "else", "enum", "event", "explicit", "extern",
            "false", "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in", "int", "interface",
            "internal", "is", "lock", "long", "namespace", "new", "null", "object", "operator", "out", "override",
            "params", "private", "protected", "public", "readonly", "ref", "return", "sbyte", "sealed", "short",
            "sizeof", "stackalloc", "static", "string", "struct", "switch", "this", "throw", "true", "try", "typeof",
            "uint", "ulong", "unchecked", "unsafe", "ushort", "using", "virtual", "void", "volatile", "while"
        },
                ["java"] = new[]
        {
            "abstract", "assert", "boolean", "break", "byte", "case", "catch", "char", "class", "const", "continue",
            "default", "do", "double", "else", "enum", "extends", "final", "finally", "float", "for", "goto", "if",
            "implements", "import", "instanceof", "int", "interface", "long", "native", "new", "null", "package",
            "private", "protected", "public", "return", "short", "static", "strictfp", "super", "switch", "synchronized",
            "this", "throw", "throws", "transient", "try", "void", "volatile", "while"
        },
                ["javascript"] = new[]
        {
            "abstract", "arguments", "await", "boolean", "break", "byte", "case", "catch", "char", "class", "const",
            "continue", "debugger", "default", "delete", "do", "double", "else", "enum", "eval", "export", "extends",
            "false", "final", "finally", "float", "for", "function", "goto", "if", "implements", "import", "in",
            "instanceof", "int", "interface", "let", "long", "native", "new", "null", "package", "private", "protected",
            "public", "return", "short", "static", "super", "switch", "synchronized", "this", "throw", "throws",
            "transient", "true", "try", "typeof", "var", "void", "volatile", "while", "with", "yield"
        }
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
                MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.Save();
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
            // For preventing unnecessary iteration of comments every time something changes in Word.
            int numComments = this.Application.ActiveDocument.Comments.Count;
            if (numComments != _prevNumComments)
            {
                Comments comments = this.Application.ActiveDocument.Comments;
                List<Word.Comment> topLevelAIComments = new List<Word.Comment>();

                foreach (Comment c in comments)
                {
                    if (c.Ancestor == null && c.Author == _model)
                    {
                        topLevelAIComments.Add(c);
                    }
                }

                foreach (Comment c in topLevelAIComments)
                {
                    if (c.Replies.Count == 0) continue;
                    if (c.Replies[c.Replies.Count].Author != _model && !_isDraftingComment)
                    {
                        _isDraftingComment = true;

                        List<ChatMessage> chatHistory = new List<ChatMessage>();
                        chatHistory.Add(Forge.SystemPrompt);

                        string documentText = (this.Application.ActiveDocument.Words.Count * 1.2 > _contextLength * 0.4) ? RAGControl.GetWordDocumentAsRAG(c.Range.Text, this.Application.ActiveDocument.Range()) : this.Application.ActiveDocument.Range().Text;
                        chatHistory.Add(new UserChatMessage($@"Document Content: ""{RAGControl.SubstringWithoutBounds(documentText, (int)(_contextLength * 0.4))}""{Environment.NewLine}Rag Context: {_ragControl.GetRAGContext(c.Range.Text, (int)(ContextLength * 0.2))}Please review the following paragraph extracted from the Document: ""{c.Range.Text}""{Environment.NewLine}Based on the previous AI comments, suggest additional specific improvements to the paragraph, focusing on clarity, coherence, structure, grammar, and overall effectiveness. Ensure that your suggestions are detailed and aimed at improving the paragraph within the context of the entire Document."));
                        chatHistory.Add(new AssistantChatMessage(c.Range.Text));

                        for (int i = 1; i <= c.Replies.Count; i++)
                        {
                            Comment reply = c.Replies[i];
                            chatHistory.Add((i % 2 == 1) ? new UserChatMessage(reply.Range.Text) : new AssistantChatMessage(reply.Range.Text));
                        }

                        ChatClient client = new ChatClient(Model, ApiKey, ClientOptions);
                        var streamingCompletion = client.CompleteChatStreamingAsync(chatHistory, null, _cancellationTokenSource.Token);
                        await Forge.AddComment(c.Replies, c.Range, streamingCompletion);
                        numComments++;

                        _isDraftingComment = false;
                    }
                }
                _prevNumComments = numComments;
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
            _modelList = Forge.GetModels(modelRetriever); // caching the response

            string defaultModel = Properties.Settings.Default.DefaultModel;
            _model = _modelList.Contains(defaultModel) ? defaultModel : modelRetriever.GetModels().Value[0].Id;

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
                    foreach (var item in ThisAddIn.EmbedModels)
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


        /// <summary>
        /// Removes various Markdown syntax elements from the provided text.
        /// </summary>
        /// <param name="markdownText">The text containing Markdown syntax.</param>
        /// <returns>The text with Markdown syntax removed.</returns>
        public static string RemoveMarkdownSyntax(string markdownText)
        {
            // Replace Environment.NewLine with \n to handle line endings consistently
            markdownText = Regex.Replace(markdownText, Environment.NewLine, "\n");

            // List of functions to remove specific Markdown syntax elements
            var removalFunctions = new Func<string, string>[]
            {
                RemoveBoldMarkdownSyntax,
                RemoveItalicMarkdownSyntax,
                RemoveUnderlineMarkdownSyntax,
                RemoveStrikeThroughMarkdownSyntax,
                RemoveBlockQuoteMarkdownSyntax,
                RemoveHeadingMarkdownSyntax,
                RemoveUnorderedListMarkdownSyntax,
                RemoveHorizontalRuleMarkdownSyntax,
                RemoveCodeBlockMarkdownSyntax,
                RemoveInlineCodeMarkdownSyntax,
                RemoveImageMarkdownSyntax,
                RemoveLinkMarkdownSyntax
            };

            // Apply each removal function to the text
            foreach (var removeFunction in removalFunctions)
            {
                markdownText = removeFunction(markdownText);
            }

            // Replace \n back with Environment.NewLine to restore original line endings
            markdownText = Regex.Replace(markdownText, "\n", Environment.NewLine);
            return markdownText;
        }

        private static string RemoveBoldMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.Bold, "$1");
        private static string RemoveItalicMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.Italic, "$1");
        private static string RemoveUnderlineMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.Underline, "$1");
        private static string RemoveStrikeThroughMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.StrikeThrough, "$1");
        private static string RemoveBlockQuoteMarkdownSyntax(string markdownText)
        {
            string pattern = RegexSyntaxFilter.BlockQuote;
            string result = markdownText;

            // Loop to remove all levels of blockquotes
            while (Regex.IsMatch(result, pattern, RegexOptions.Multiline))
            {
                result = Regex.Replace(result, pattern, "$2", RegexOptions.Multiline);
            }

            return result;
        }
        private static string RemoveHeadingMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.Heading, "$2", RegexOptions.Multiline);
        private static string RemoveUnorderedListMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.UnorderedList, "$1", RegexOptions.Multiline);
        private static string RemoveHorizontalRuleMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.HorizontalRule, string.Empty, RegexOptions.Multiline);
        private static string RemoveInlineCodeMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.InlineCode, "$1", RegexOptions.Singleline);
        private static string RemoveCodeBlockMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.CodeBlock, "$2", RegexOptions.Singleline);
        private static string RemoveLinkMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.Link, "$1");
        private static string RemoveBoldItalicMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.BoldItalic, "$1$2", RegexOptions.Multiline);
        private static string RemoveImageMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.Image, string.Empty, RegexOptions.Multiline);

        /// <summary>
        /// Removes all Markdown syntax from the given text except for the specified syntax type.
        /// </summary>
        /// <param name="num">The Markdown syntax type to retain.</param>
        /// <param name="markdownText">The text from which to remove Markdown syntax.</param>
        /// <returns>The text with the specified Markdown syntax retained and all other Markdown syntax removed.</returns>
        private static string RemoveAllMarkdownSyntaxExcept(RegexSyntaxFilter.Number num, string markdownText)
        {
            // Dictionary mapping each Markdown syntax type to its corresponding removal function
            var syntaxRemovers = new Dictionary<RegexSyntaxFilter.Number, Action<string>>
            {
                { RegexSyntaxFilter.Number.Bold, text => markdownText = RemoveBoldMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.Italic, text => markdownText = RemoveItalicMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.Underline, text => markdownText = RemoveUnderlineMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.StrikeThrough, text => markdownText = RemoveStrikeThroughMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.BlockQuote, text => markdownText = RemoveBlockQuoteMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.Heading, text => markdownText = RemoveHeadingMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.UnorderedList, text => markdownText = RemoveUnorderedListMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.HorizontalRule, text => markdownText = RemoveHorizontalRuleMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.CodeBlock, text => markdownText = RemoveCodeBlockMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.InlineCode, text => markdownText = RemoveInlineCodeMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.BoldItalic, text => markdownText = RemoveBoldItalicMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.Image, text => markdownText = RemoveImageMarkdownSyntax(text) },
                { RegexSyntaxFilter.Number.Link, text => markdownText = RemoveLinkMarkdownSyntax(text) },
            };

            // Iterate through the dictionary and apply all removals except the one specified by 'num'
            foreach (var remover in syntaxRemovers)
            {
                if (remover.Key != num)
                {
                    remover.Value(markdownText);
                }
            }

            return markdownText;
        }


        /// <summary>
        /// Applies all Markdown formatting types to a specified range in a Word document.
        /// </summary>
        /// <param name="commentRange">The range in the Word document to apply formatting to.</param>
        /// <param name="rawMarkdown">The raw Markdown text to be processed.</param>
        public static void ApplyAllMarkdownFormatting(Word.Range commentRange, string rawMarkdown)
        {
            // Word treats \r\n as a single character but .Length returns 2!
            rawMarkdown = Regex.Replace(rawMarkdown, "\r", string.Empty);

            // Define all the Markdown formatting types to be applied
            var formattingTypes = new[]
            {
                RegexSyntaxFilter.Number.Bold,
                RegexSyntaxFilter.Number.Italic,
                RegexSyntaxFilter.Number.Underline,
                RegexSyntaxFilter.Number.StrikeThrough,
                RegexSyntaxFilter.Number.BlockQuote,
                RegexSyntaxFilter.Number.Heading,
                RegexSyntaxFilter.Number.UnorderedList,
                RegexSyntaxFilter.Number.HorizontalRule,
                RegexSyntaxFilter.Number.InlineCode,
                RegexSyntaxFilter.Number.CodeBlock,
                RegexSyntaxFilter.Number.Image,
                RegexSyntaxFilter.Number.Link,
                RegexSyntaxFilter.Number.OrderedList
            };

            // Apply each Markdown formatting type to the specified range
            foreach (var formattingType in formattingTypes)
            {
                ApplyMarkdownFormatting(commentRange, rawMarkdown, formattingType);
            }
        }


        // TODO: Support nested ordered/unordered list
        /// <summary>
        /// Applies Markdown formatting to a specified range in a Word document.
        /// </summary>
        /// <param name="commentRange">The range in the Word document to apply formatting to.</param>
        /// <param name="fullMarkdownText">The full Markdown text to be processed.</param>
        /// <param name="formatType">The type of Markdown formatting to apply.</param>
        private static void ApplyMarkdownFormatting(Word.Range commentRange, string fullMarkdownText, RegexSyntaxFilter.Number formatType)
        {
            // Determine the appropriate regex based on the format type
            Regex regex = formatType switch
            {
                RegexSyntaxFilter.Number.Bold => new Regex(RegexSyntaxFilter.Bold),
                RegexSyntaxFilter.Number.Italic => new Regex(RegexSyntaxFilter.Italic),
                RegexSyntaxFilter.Number.Underline => new Regex(RegexSyntaxFilter.Underline),
                RegexSyntaxFilter.Number.StrikeThrough => new Regex(RegexSyntaxFilter.StrikeThrough),
                RegexSyntaxFilter.Number.BlockQuote => new Regex(RegexSyntaxFilter.BlockQuote, RegexOptions.Multiline),
                RegexSyntaxFilter.Number.Heading => new Regex(RegexSyntaxFilter.Heading, RegexOptions.Multiline),
                RegexSyntaxFilter.Number.UnorderedList => new Regex(RegexSyntaxFilter.UnorderedList, RegexOptions.Multiline),
                RegexSyntaxFilter.Number.OrderedList => new Regex(RegexSyntaxFilter.OrderedList, RegexOptions.Multiline),
                RegexSyntaxFilter.Number.HorizontalRule => new Regex(RegexSyntaxFilter.HorizontalRule, RegexOptions.Multiline),
                RegexSyntaxFilter.Number.InlineCode => new Regex(RegexSyntaxFilter.InlineCode, RegexOptions.Multiline),
                RegexSyntaxFilter.Number.CodeBlock => new Regex(RegexSyntaxFilter.CodeBlock, RegexOptions.Singleline),
                RegexSyntaxFilter.Number.Link => new Regex(RegexSyntaxFilter.Link),
                RegexSyntaxFilter.Number.BoldItalic => new Regex(RegexSyntaxFilter.BoldItalic),
                RegexSyntaxFilter.Number.Image => new Regex(RegexSyntaxFilter.Image),
                _ => throw new ArgumentOutOfRangeException("Unknown format type for Markdown processing!"),
            };

            // Remove all Markdown syntax except the specified format type
            string partialMarkdownText = RemoveAllMarkdownSyntaxExcept(formatType, fullMarkdownText);
            MatchCollection matches = regex.Matches(partialMarkdownText);

            int searchIndex = 0;
            int offset = 0;
            foreach (Match match in matches)
            {
                string textToFormat = match.Value;
                string insideContent = match.Groups[1].Value;
                searchIndex = commentRange.Start + partialMarkdownText.IndexOf(match.Value, searchIndex);
                int length = textToFormat.Length;

                Word.Range formatRange = commentRange.Duplicate;
                switch (formatType)
                {
                    case RegexSyntaxFilter.Number.Bold:
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 1, ref offset, 4);
                        break;
                    case RegexSyntaxFilter.Number.Italic:
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 2, ref offset, 2);
                        break;
                    case RegexSyntaxFilter.Number.Underline:
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 3, ref offset, 4);
                        break;
                    case RegexSyntaxFilter.Number.StrikeThrough:
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 4, ref offset, 4);
                        break;
                    case RegexSyntaxFilter.Number.BlockQuote:
                        int level = match.Groups[1].Value.Length; // Determine the level of nesting
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 5, ref offset, level + 1, level);
                        break;
                    case RegexSyntaxFilter.Number.Heading:
                        ApplyHeadingFormatting(formatRange, searchIndex, offset, match, insideContent, ref offset);
                        break;
                    case RegexSyntaxFilter.Number.UnorderedList:
                        ApplyFormatting(formatRange, searchIndex, offset, match.Groups[2].Length, 6, ref offset, 2);
                        break;
                    case RegexSyntaxFilter.Number.OrderedList:
                        ApplyFormatting(formatRange, searchIndex, offset, textToFormat.Length - 3, 9, ref offset, 3);
                        break;
                    case RegexSyntaxFilter.Number.HorizontalRule:
                        ApplyFormatting(formatRange, searchIndex, offset, textToFormat.Length, 7, ref offset, length);
                        break;
                    case RegexSyntaxFilter.Number.InlineCode:
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 8, ref offset, 2);
                        break;
                    case RegexSyntaxFilter.Number.CodeBlock:
                        ApplyCodeBlockFormatting(formatRange, searchIndex, offset, match, ref offset);
                        break;
                    case RegexSyntaxFilter.Number.Link:
                        ApplyLinkFormatting(formatRange, searchIndex, offset, match, ref offset);
                        break;
                    case RegexSyntaxFilter.Number.BoldItalic:
                        ApplyFormatting(formatRange, searchIndex, offset, insideContent.Length, 10, ref offset, 6);
                        break;
                    case RegexSyntaxFilter.Number.Image:
                        string imageUrl = match.Groups[2].Value;
                        ApplyImageFormatting(formatRange, imageUrl);
                        offset += length; // Adjust the offset to account for the removed text
                        break;
                }
                searchIndex += length;
            }
        }

        /// <summary>
        /// Applies various formatting styles to a specified range in a Word document.
        /// </summary>
        /// <param name="formatRange">The range in the Word document to format.</param>
        /// <param name="searchIndex">The starting index of the match in the document.</param>
        /// <param name="offset">The offset to adjust the range.</param>
        /// <param name="length">The length of the text to format.</param>
        /// <param name="formatType">The type of formatting to apply (1: Bold, 2: Italic, 3: Underline, 4: Strikethrough, 5: Blockquote, 6: Bullet list, 7: Horizontal line, 8: Courier New font, 9: Numbered list).</param>
        /// <param name="offsetIncrement">The increment to adjust the offset.</param>
        /// <param name="offsetValue">The value to add to the offset increment.</param>
        private static void ApplyFormatting(Word.Range formatRange, int searchIndex, int offset, int length, int formatType, ref int offsetIncrement, int offsetValue, int level = 1)
        {
            formatRange.SetRange(searchIndex - offset, searchIndex - offset + length);
            switch (formatType)
            {
                case 1:
                    formatRange.Font.Bold = 1;
                    break;
                case 2:
                    formatRange.Font.Italic = 1;
                    break;
                case 3:
                    formatRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    break;
                case 4:
                    formatRange.Font.StrikeThrough = 1;
                    break;
                case 5:
                    FormatAsBlockquote(formatRange, level);
                    break;
                case 6:
                    formatRange.ListFormat.ApplyBulletDefault();
                    break;
                case 7:
                    formatRange.InlineShapes.AddHorizontalLineStandard();
                    break;
                case 8:
                    formatRange.Font.Name = "Courier New";
                    break;
                case 9:
                    formatRange.ListFormat.ApplyNumberDefault();
                    break;
                case 10:
                    formatRange.Font.Bold = 1;
                    formatRange.Font.Italic = 1;
                    break;
            }
            offsetIncrement += offsetValue;
        }

        private static void ApplyImageFormatting(Word.Range formatRange, string imageUrl)
        {
            using (var webClient = new System.Net.WebClient())
            {
                byte[] imageBytes = webClient.DownloadData(imageUrl);
                string tempFilePath = System.IO.Path.GetTempFileName();
                System.IO.File.WriteAllBytes(tempFilePath, imageBytes);
                formatRange.InlineShapes.AddPicture(tempFilePath);
                System.IO.File.Delete(tempFilePath); // Clean up the temporary file
            }
        }

        /// <summary>
        /// Applies heading formatting to a specified range in a Word document.
        /// </summary>
        /// <param name="formatRange">The range in the Word document to format.</param>
        /// <param name="searchIndex">The starting index of the match in the document.</param>
        /// <param name="offset">The offset to adjust the range.</param>
        /// <param name="match">The match object containing the matched text.</param>
        /// <param name="insideContent">The content inside the markdown heading.</param>
        /// <param name="offsetIncrement">The increment to adjust the offset.</param>
        /// <exception cref="ArgumentException">Thrown when the number of # (markdown heading) is not valid.</exception>
        private static void ApplyHeadingFormatting(Word.Range formatRange, int searchIndex, int offset, Match match, string insideContent, ref int offsetIncrement)
        {
            formatRange.SetRange(searchIndex - offset, searchIndex - offset + match.Groups[2].Length);
            if (insideContent.Length < 1 || insideContent.Length > 6)
                throw new ArgumentException("The number of # (markdown heading) is not valid!");
            formatRange.set_Style($"Heading {insideContent.Length}");
            offsetIncrement += insideContent.Length + 1;
        }

        /// <summary>
        /// Applies code block formatting to a specified range in a Word document.
        /// </summary>
        /// <param name="formatRange">The range in the Word document to format.</param>
        /// <param name="searchIndex">The starting index of the match in the document.</param>
        /// <param name="offset">The offset to adjust the range.</param>
        /// <param name="match">The match object containing the matched text.</param>
        /// <param name="offsetIncrement">The increment to adjust the offset.</param>
        private static void ApplyCodeBlockFormatting(Word.Range formatRange, int searchIndex, int offset, Match match, ref int offsetIncrement)
        {
            string language = match.Groups[1].Value;
            string code = match.Groups[2].Value;
            ApplySyntaxHighlighting(formatRange, language, code);
            offsetIncrement += 6;
        }

        /// <summary>
        /// Applies link formatting to a specified range in a Word document.
        /// </summary>
        /// <param name="formatRange">The range in the Word document to format.</param>
        /// <param name="searchIndex">The starting index of the match in the document.</param>
        /// <param name="offset">The offset to adjust the range.</param>
        /// <param name="match">The match object containing the matched text.</param>
        /// <param name="offsetIncrement">The increment to adjust the offset.</param>
        private static void ApplyLinkFormatting(Word.Range formatRange, int searchIndex, int offset, Match match, ref int offsetIncrement)
        {
            string linkText = match.Groups[1].Value;
            string linkUrl = match.Groups[2].Value;
            formatRange.SetRange(searchIndex - offset, searchIndex - offset + linkText.Length);
            formatRange.Hyperlinks.Add(formatRange, linkUrl, Type.Missing, Type.Missing, linkText);
            offsetIncrement += 4 + linkUrl.Length;
        }

        private static void FormatAsBlockquote(Word.Range range, int level)
        {
            Word.ParagraphFormat format = range.ParagraphFormat;
            format.LeftIndent = range.Application.InchesToPoints(0.5f * level); // Indent by 0.5 inches per level
            format.SpaceBefore = 12; // Space before paragraph
            format.SpaceAfter = 12; // Space after paragraph
            format.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            format.Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
            format.Borders[Word.WdBorderType.wdBorderLeft].Color = Word.WdColor.wdColorGray25;
        }

        /// <summary>
        /// Applies syntax highlighting to the specified range in a Word document based on the provided programming language and code.
        /// </summary>
        /// <param name="range">The Word range where the syntax highlighting will be applied.</param>
        /// <param name="language">The programming language of the code.</param>
        /// <param name="code">The code to be highlighted.</param>
        private static void ApplySyntaxHighlighting(Word.Range range, string language, string code)
        {
            // Retrieve the keywords for the specified language
            var keywords = GetKeywordsForLanguage(language);

            // Create a regex pattern to match any of the keywords
            var regex = new Regex(@"\b(" + string.Join("|", keywords) + @")\b(?=\[\])?");

            // Iterate over each match in the code
            foreach (Match match in regex.Matches(code))
            {
                // Duplicate the range and set the range to the match's location
                var keywordRange = range.Duplicate;
                keywordRange.SetRange(range.Start + match.Index, range.Start + match.Index + match.Length);

                // Apply blue color to the matched keyword
                keywordRange.Font.Color = Word.WdColor.wdColorBlue;
            }
        }
        private static string[] GetKeywordsForLanguage(string language)
        {
            return Keywords.TryGetValue(language.ToLower(), out var keywords) ? keywords : Array.Empty<string>();
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

    public class RegexSyntaxFilter
    {
        public enum Number
        {
            Bold, Italic, Underline, StrikeThrough, BlockQuote, HorizontalRule, InlineCode, CodeBlock, Link, Heading, UnorderedList, OrderedList, BoldItalic, Image
        }

        public const string Bold = @"\*\*(.*?)\*\*|__(.*?)__";
        public const string Italic = @"(?<!\*)\*(?!\*)(.*?)\*(?!\*)|(?<!_)_(?!_)(.*?)_(?!_)";
        public const string Underline = @"__(.*?)__";
        public const string StrikeThrough = @"~~(.*?)~~";
        public const string BoldItalic = @"\*\*\*(.*?)\*\*\*|___(.*?)___";

        public const string HorizontalRule = @"^(?:-{3,}|_{3,}|\*{3,})\s*$";
        public const string Heading = @"^(#{1,6})\s*(.+)$";
        public const string BlockQuote = @"^(>+)\s?(.*)";
        public const string UnorderedList = @"^[\*\-\+]\s+(.*)";
        public const string OrderedList = @"^\d+\.\s[^\n]*";

        public const string Link = @"(?<!\!)\[(.*?)\]\((.*?)\)";
        public const string InlineCode = @"(?<!`)`([^`]*)`(?!`)";
        public const string CodeBlock = @"```(\w+)?\s*([\s\S]*?)```";
        public const string Image = @"!\[(.*?)\]\((.*?)\)";
    }
}