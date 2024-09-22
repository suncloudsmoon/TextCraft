using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    internal class CommonUtils
    {
        public static void DisplayError(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void DisplayWarning(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void DisplayInformation(Exception ex)
        {
            MessageBox.Show(ex.Message, ex.GetType().Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static bool GetInternetAccessPermission(string url)
        {
            var result = MessageBox.Show($"Do you want to allow TextCraft to access the following internet resource?{Environment.NewLine}{url}", "Internet Access", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            return result == DialogResult.Yes;
        }

        public static Word.Application GetApplication()
        {
            return Globals.ThisAddIn.Application;
        }

        public static Document GetActiveDocument()
        {
            return Globals.ThisAddIn.Application.ActiveDocument;
        }

        public static Comments GetComments()
        {
            return GetActiveDocument().Comments;
        }

        public static Range GetSelectionRange()
        {
            return GetApplication().Selection.Range;
        }
        public static int GetWordPageCount()
        {
            int pageCount = GetActiveDocument().ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
            return pageCount;
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
        public static int CharToTokenCount(int charCount)
        {
            return charCount / 4; // https://platform.openai.com/tokenizer
        }

        public static void GetEnvironmentVariableIfAvailable(ref string dest, string variable)
        {
            var key = Environment.GetEnvironmentVariable(variable);
            if (key != null)
                dest = key;
        }
    }
}
