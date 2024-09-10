using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    internal class WordMarkdown
    {
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
        private static string RemoveAlternateHeadingMarkdownSyntax(string markdownText) => Regex.Replace(markdownText, RegexSyntaxFilter.AlternateHeading, "$1", RegexOptions.Multiline);

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
                RemoveAlternateHeadingMarkdownSyntax,
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
                { RegexSyntaxFilter.Number.AlternateHeading, text => markdownText = RemoveAlternateHeadingMarkdownSyntax(text) },
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
                RegexSyntaxFilter.Number.AlternateHeading,
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
                RegexSyntaxFilter.Number.AlternateHeading => new Regex(RegexSyntaxFilter.AlternateHeading, RegexOptions.Multiline),
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
                    case RegexSyntaxFilter.Number.AlternateHeading:
                        ApplyAlternateHeadingFormatting(formatRange, searchIndex, offset, match, insideContent, ref offset);
                        break;
                }
                searchIndex += length;
            }
        }

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

        private static void ApplyAlternateHeadingFormatting(Word.Range formatRange, int searchIndex, int offset, Match match, string insideContent, ref int offsetIncrement)
        {
            formatRange.SetRange(searchIndex - offset, searchIndex - offset + match.Groups[1].Length);
            formatRange.set_Style("Heading 1");
            offsetIncrement += match.Groups[2].Length + 1; // accounting for '\n'
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

         private static void ApplyHeadingFormatting(Word.Range formatRange, int searchIndex, int offset, Match match, string insideContent, ref int offsetIncrement)
        {
            formatRange.SetRange(searchIndex - offset, searchIndex - offset + match.Groups[2].Length);
            if (insideContent.Length < 1 || insideContent.Length > 6)
                throw new ArgumentException("The number of # (markdown heading) is not valid!");
            formatRange.set_Style($"Heading {insideContent.Length}");
            offsetIncrement += insideContent.Length + 1;
        }

        private static void ApplyCodeBlockFormatting(Word.Range formatRange, int searchIndex, int offset, Match match, ref int offsetIncrement)
        {
            string language = match.Groups[1].Value;
            string code = match.Groups[2].Value;
            ApplySyntaxHighlighting(formatRange, language, code);
            offsetIncrement += 6;
        }

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
    }

    public class RegexSyntaxFilter
    {
        public enum Number
        {
            Bold, Italic, Underline, StrikeThrough, BlockQuote, HorizontalRule, InlineCode, CodeBlock, Link, Heading, UnorderedList, OrderedList, BoldItalic, Image, AlternateHeading
        }

        public const string Bold = @"\*\*(.*?)\*\*|__(.*?)__";
        // TODO: fix "_italic_" behavior
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

        public const string AlternateHeading = @"^(?!.*(?:-{3,}|_{3,}|\*{3,})\s*$)(.*)\n(=+)$";
    }
}
