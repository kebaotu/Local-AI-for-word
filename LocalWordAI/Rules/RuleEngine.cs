using LocalWordAI.Models;
using System.Collections.Generic;

namespace LocalWordAI.Rules
{
    public class RuleEngine
    {
        private readonly SpacingChecker _spacing = new SpacingChecker();
        private readonly PunctuationChecker _punctuation = new PunctuationChecker();
        private readonly RepeatedWordChecker _repeated = new RepeatedWordChecker();
        private readonly DateLogicChecker _date = new DateLogicChecker();
        private readonly NumberingChecker _numbering = new NumberingChecker();
        private readonly CrossReferenceChecker _crossRef = new CrossReferenceChecker();
        private readonly PlaceholderChecker _placeholder = new PlaceholderChecker();

        public List<Suggestion> RunAll(string text)
        {
            var all = new List<Suggestion>();
            if (string.IsNullOrEmpty(text)) return all;

            all.AddRange(_spacing.Check(text));
            all.AddRange(_punctuation.Check(text));
            all.AddRange(_repeated.Check(text));
            all.AddRange(_date.Check(text));
            all.AddRange(_numbering.Check(text));
            all.AddRange(_crossRef.Check(text));
            all.AddRange(_placeholder.Check(text));

            return all;
        }

        public List<Suggestion> RunQuick(string text)
        {
            var all = new List<Suggestion>();
            all.AddRange(_spacing.Check(text));
            all.AddRange(_punctuation.Check(text));
            all.AddRange(_repeated.Check(text));
            all.AddRange(_placeholder.Check(text));
            return all;
        }
    }
}
