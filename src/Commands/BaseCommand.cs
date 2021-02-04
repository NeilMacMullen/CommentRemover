using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text.RegularExpressions;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;

namespace CommentRemover
{
    internal abstract class BaseCommand<T> where T : BaseCommand<T>, new()
    {
        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package)
        {
            Instance = new T
            {
                DTE = await package.GetServiceAsync(typeof(DTE)) as DTE2,
                CommandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService
            };

            Instance.SetupCommands();
        }

        protected abstract void SetupCommands();

        protected void RegisterCommand(Guid commandGuid, int commandId)
        {
            var id = new CommandID(commandGuid, commandId);
            var command = new OleMenuCommand(Callback, id);
            CommandService.AddCommand(command);
        }

        protected DTE2 DTE { get; private set; }

        private OleMenuCommandService CommandService { get; set; }

        public static T Instance { get; protected set; }

        private void Callback(object sender, EventArgs e)
        {
            var button = (OleMenuCommand)sender;
            Execute(button);
        }

        protected abstract void Execute(OleMenuCommand button);

        protected static IEnumerable<IMappingSpan> GetClassificationSpans(IWpfTextView view, string classificationName)
        {
            if (view == null)
                return Enumerable.Empty<IMappingSpan>();

            var componentModel = ProjectHelpers.GetComponentModel();
            var service = componentModel.GetService<IBufferTagAggregatorFactoryService>();
            var classifier = service.CreateTagAggregator<IClassificationTag>(view.TextBuffer);
            var snapshot = new SnapshotSpan(view.TextBuffer.CurrentSnapshot, 0, view.TextBuffer.CurrentSnapshot.Length);

            return from s in classifier.GetTags(snapshot).Reverse()
                   where s.Tag.ClassificationType.Classification.IndexOf(classificationName, StringComparison.OrdinalIgnoreCase) > -1
                   select s.Span;
        }

        protected static bool IsLineEmpty(ITextSnapshotLine line)
        {
            var text = line.GetText().Trim();

            return (string.IsNullOrWhiteSpace(text)
                   || text == "<!--"
                   || text == "-->"
                   || text == "<%%>"
                   || text == "<%"
                   || text == "%>"
                   || Regex.IsMatch(text, @"<!--(\s+)?-->"));
        }

        protected static bool IsEmptyXmlReturns(ITextSnapshotLine line)
        {
            var text = line.GetText().Trim();
            var contentType = line.Snapshot.TextBuffer.ContentType;

            if ((contentType.IsOfType("CSharp") || contentType.IsOfType("FSharp")) && text.StartsWith("///"))
            {
                return Regex.IsMatch(text, @"(\s+)?/// <returns>(\s+)?</returns>(\s+)?");
            }

            if (contentType.IsOfType("Basic") && text.StartsWith("'''"))
            {
                return Regex.IsMatch(text, @"(\s+)?''' <returns>(\s+)?</returns>(\s+)?");
            }

            return false;
        }

        protected static bool IsEmptyXmlRemarks(ITextSnapshotLine line)
        {
            var text = line.GetText().Trim();
            var contentType = line.Snapshot.TextBuffer.ContentType;

            if ((contentType.IsOfType("CSharp") || contentType.IsOfType("FSharp")) && text.StartsWith("///"))
            {
                return Regex.IsMatch(text, @"(\s+)?/// <remarks>(\s+)?</remarks>(\s+)?");
            }

            if (contentType.IsOfType("Basic") && text.StartsWith("'''"))
            {
                return Regex.IsMatch(text, @"(\s+)?''' <remarks>(\s+)?</remarks>(\s+)?");
            }

            return false;
        }

        protected static bool IsXmlDocComment(ITextSnapshotLine line)
        {
            string text = line.GetText().Trim();
            var contentType = line.Snapshot.TextBuffer.ContentType;

            if (contentType.IsOfType("CSharp") && text.StartsWith("///"))
                return true;

            if (contentType.IsOfType("FSharp") && text.StartsWith("///"))
                return true;

            if (contentType.IsOfType("Basic") && text.StartsWith("'''"))
                return true;

            return false;
        }
    }
}
