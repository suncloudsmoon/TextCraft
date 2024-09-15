using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using OpenAI.Chat;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    internal class CommentHandler
    {
        private static int _prevNumComments = 0;
        private static bool _isDraftingComment = false;

        public static async void Document_CommentsEventHandler(Word.Selection selection)
        {
            try
            {
                // For preventing unnecessary iteration of comments every time something changes in Word.
                int numComments = Globals.ThisAddIn.Application.ActiveDocument.Comments.Count;
                if (numComments == _prevNumComments) return;

                var topLevelAIComments = GetTopLevelAIComments(CommonUtils.GetComments());
                foreach (Comment c in topLevelAIComments)
                {
                    if (c.Replies.Count == 0) continue;
                    if (c.Replies[c.Replies.Count].Author != ThisAddIn.Model && !_isDraftingComment)
                    {
                        _isDraftingComment = true;

                        List<ChatMessage> chatHistory = new List<ChatMessage>()
                            {
                                new UserChatMessage($@"Please review the following paragraph extracted from the Document: ""{CommonUtils.SubstringTokens(c.Range.Text, (int)(ThisAddIn.ContextLength * 0.2))}"""),
                                new UserChatMessage($@"Based on the previous AI comments, suggest additional specific improvements to the paragraph, focusing on clarity, coherence, structure, grammar, and overall effectiveness. Ensure that your suggestions are detailed and aimed at improving the paragraph within the context of the entire Document.")
                            };
                        for (int i = 1; i <= c.Replies.Count; i++)
                        {
                            Comment reply = c.Replies[i];
                            chatHistory.Add((i % 2 == 1) ? new UserChatMessage(reply.Range.Text) : new AssistantChatMessage(reply.Range.Text));
                        }
                        await AddComment(
                            c.Replies,
                            c.Range,
                            RAGControl.AskQuestion(Forge.CommentSystemPrompt, chatHistory, CommonUtils.GetActiveDocument().Range())
                        );
                        numComments++;

                        _isDraftingComment = false;
                    }
                }
                _prevNumComments = numComments;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private static IEnumerable<Comment> GetTopLevelAIComments(Comments comments)
        {
            List<Comment> topLevelAIComments = new List<Comment>();
            foreach (Comment c in comments)
                if (c.Ancestor == null && c.Author == ThisAddIn.Model)
                    topLevelAIComments.Add(c);
            return topLevelAIComments;
        }

        public static async Task AddComment(Word.Comments comments, Word.Range range, AsyncCollectionResult<StreamingChatCompletionUpdate> streamingContent)
        {
            Word.Comment c = comments.Add(range, string.Empty);
            c.Author = ThisAddIn.Model;
            Word.Range commentRange = c.Range.Duplicate; // Duplicate the range to work with

            StringBuilder comment = new StringBuilder();
            // Run the comment generation in a background thread
            await Task.Run(async () =>
            {
                Forge.CancelButtonVisibility(true);
                await foreach (var update in streamingContent.WithCancellation(ThisAddIn.CancellationTokenSource.Token))
                {
                    if (ThisAddIn.CancellationTokenSource.IsCancellationRequested)
                        break;
                    foreach (var content in update.ContentUpdate)
                    {
                        commentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd); // Move to the end of the range
                        commentRange.Text = content.Text; // Append new text
                        commentRange = c.Range.Duplicate; // Update the range to include the new text
                        comment.Append(content.Text);
                    }
                }
                Forge.CancelButtonVisibility(false);
                c.Range.Text = WordMarkdown.RemoveMarkdownSyntax(comment.ToString());
            });
        }
    }
}
