using LocalWordAI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace LocalWordAI.WordIntegration
{
    public class TrackChangeService
    {
        private readonly Word.Application _app;

        public TrackChangeService(Word.Application app)
        {
            _app = app;
        }

        public bool ApplySuggestion(Word.Document doc, Suggestion suggestion)
        {
            if (suggestion == null || doc == null) return false;
            if (string.IsNullOrEmpty(suggestion.RangeText)) return false;

            bool oldTrack = doc.TrackRevisions;
            bool oldScreen = _app.ScreenUpdating;
            try
            {
                doc.TrackRevisions = true;
                _app.ScreenUpdating = false;

                var range = FindRange(doc, suggestion.RangeText, suggestion.RangeStart);
                if (range == null) return false;

                switch (suggestion.Action)
                {
                    case SuggestionAction.Replace:
                        if (!string.IsNullOrEmpty(suggestion.ReplacementText))
                            range.Text = suggestion.ReplacementText;
                        break;
                    case SuggestionAction.Delete:
                        range.Text = "";
                        break;
                    case SuggestionAction.Insert:
                        range.InsertAfter(suggestion.ReplacementText ?? "");
                        break;
                    default:
                        return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ApplySuggestion error: " + ex.Message);
                return false;
            }
            finally
            {
                doc.TrackRevisions = oldTrack;
                _app.ScreenUpdating = oldScreen;
            }
        }

        public int ApplyMultiple(Word.Document doc, List<Suggestion> suggestions)
        {
            // Apply from end to beginning to avoid range shift
            var sorted = suggestions
                .Where(s => s.Action == SuggestionAction.Replace || s.Action == SuggestionAction.Delete)
                .OrderByDescending(s => s.RangeStart)
                .ToList();

            bool oldTrack = doc.TrackRevisions;
            bool oldScreen = _app.ScreenUpdating;
            int count = 0;

            try
            {
                doc.TrackRevisions = true;
                _app.ScreenUpdating = false;

                foreach (var s in sorted)
                {
                    var range = FindRange(doc, s.RangeText, s.RangeStart);
                    if (range == null) continue;

                    if (s.Action == SuggestionAction.Replace && !string.IsNullOrEmpty(s.ReplacementText))
                        range.Text = s.ReplacementText;
                    else if (s.Action == SuggestionAction.Delete)
                        range.Text = "";

                    count++;
                }
            }
            finally
            {
                doc.TrackRevisions = oldTrack;
                _app.ScreenUpdating = oldScreen;
            }
            return count;
        }

        private Word.Range FindRange(Word.Document doc, string searchText, int hintStart = -1)
        {
            // Exact match via Find
            try
            {
                var find = doc.Content.Duplicate;
                find.Find.ClearFormatting();
                find.Find.Text = searchText;
                find.Find.Forward = true;
                find.Find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Find.MatchCase = true;
                if (find.Find.Execute())
                    return find;
            }
            catch { }

            // Fallback: use hint position
            if (hintStart >= 0)
            {
                try
                {
                    var range = doc.Range(hintStart, hintStart + searchText.Length);
                    if (range.Text == searchText) return range;
                }
                catch { }
            }

            return null;
        }
    }
}
