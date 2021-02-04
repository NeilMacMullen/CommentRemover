using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace CommentRemover
{
    internal sealed class RemoveEmptyXmlRemarksCommand : BaseCommand<RemoveEmptyXmlRemarksCommand>
    {
        protected override void SetupCommands()
        {
            RegisterCommand(PackageGuids.guidPackageCmdSet, PackageIds.RemoveEmptyXmlRemarks);
        }

        protected override void Execute(OleMenuCommand button)
        {
            var view = ProjectHelpers.GetCurentTextView();
            var mappingSpans = GetClassificationSpans(view, "comment");

            try
            {
                DTE.UndoContext.Open(button.Text);
                RemoveEmptyXmlRemarks(view, mappingSpans);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex);
            }
            finally
            {
                DTE.UndoContext.Close();
            }
        }

        private void RemoveEmptyXmlRemarks(IWpfTextView view, IEnumerable<IMappingSpan> mappingSpans)
        {
            var affectedLines = new List<int>();

            using (var edit = view.TextBuffer.CreateEdit())
            {
                foreach (var mappingSpan in mappingSpans)
                {
                    var start = mappingSpan.Start.GetPoint(view.TextBuffer, PositionAffinity.Predecessor).Value;
                    var end = mappingSpan.End.GetPoint(view.TextBuffer, PositionAffinity.Successor).Value;

                    var span = new Span(start, end - start);
                    var line = view.TextBuffer.CurrentSnapshot.Lines.First(l => l.Extent.IntersectsWith(span));

                    if (IsEmptyXmlRemarks(line))
                    {
                        edit.Delete(span);

                        if (!affectedLines.Contains(line.LineNumber))
                        {
                            affectedLines.Add(line.LineNumber);
                        }
                    }
                }

                edit.Apply();
            }

            using (var edit = view.TextBuffer.CreateEdit())
            {
                foreach (var lineNumber in affectedLines)
                {
                    var line = view.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineNumber);

                    if (IsLineEmpty(line))
                    {
                        edit.Delete(line.Start, line.LengthIncludingLineBreak);
                    }
                }

                edit.Apply();
            }
        }
    }
}