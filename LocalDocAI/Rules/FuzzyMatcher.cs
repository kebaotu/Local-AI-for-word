using System;
using System.Collections.Generic;

namespace LocalDocAI.Rules
{
    // Simple Levenshtein-based fuzzy matcher (no external library dependency for basic use)
    public static class FuzzyMatcher
    {
        public static int Levenshtein(string a, string b)
        {
            if (string.IsNullOrEmpty(a)) return b?.Length ?? 0;
            if (string.IsNullOrEmpty(b)) return a.Length;

            var d = new int[a.Length + 1, b.Length + 1];
            for (int i = 0; i <= a.Length; i++) d[i, 0] = i;
            for (int j = 0; j <= b.Length; j++) d[0, j] = j;

            for (int i = 1; i <= a.Length; i++)
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }

            return d[a.Length, b.Length];
        }

        public static double Ratio(string a, string b)
        {
            if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return 1.0;
            int maxLen = Math.Max(a?.Length ?? 0, b?.Length ?? 0);
            if (maxLen == 0) return 1.0;
            return 1.0 - (double)Levenshtein(a, b) / maxLen;
        }

        public static string BestMatch(string query, IEnumerable<string> candidates, double threshold = 0.7)
        {
            string best = null;
            double bestScore = 0;
            foreach (var c in candidates)
            {
                double score = Ratio(query, c);
                if (score > bestScore && score >= threshold)
                {
                    bestScore = score;
                    best = c;
                }
            }
            return best;
        }
    }
}
