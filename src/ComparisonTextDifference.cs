using System;
using System.Collections.Generic;
using System.Drawing;

namespace FormulaExplorer
{
    internal struct ComparisonTextSpan
    {
        internal readonly int Start, Length;
        internal ComparisonTextSpan(int start, int length) { Start = start; Length = length; }
    }

    internal sealed class ComparisonTextDifference
    {
        const int MaxDistance = 2000;
        internal readonly List<ComparisonTextSpan> Left = new List<ComparisonTextSpan>();
        internal readonly List<ComparisonTextSpan> Right = new List<ComparisonTextSpan>();

        internal static ComparisonTextDifference Find(string left, string right, bool whole, int maxDistance = MaxDistance)
        {
            var difference = new ComparisonTextDifference();
            if (left == null || right == null) return difference;
            if (!whole && String.Equals(left, right, StringComparison.Ordinal)) return difference;

            List<Point> matches = null;
            if (!whole && left.Length != 0 && right.Length != 0 && Math.Abs(left.Length - right.Length) <= maxDistance)
            {
                var words = new Dictionary<string, int>(StringComparer.Ordinal);
                int nextSymbol = Char.MaxValue + 1;
                int[] a = Symbols(left, words, ref nextSymbol), b = Symbols(right, words, ref nextSymbol);
                matches = ComparisonAlignment.Match(a.Length, b.Length, (x, y) => a[x] == b[y], maxDistance);
            }
            int startLeft = 0, startRight = 0;
            if (matches != null)
                foreach (Point match in matches)
                {
                    AddGap(difference.Left, startLeft, match.X);
                    AddGap(difference.Right, startRight, match.Y);
                    startLeft = match.X + 1;
                    startRight = match.Y + 1;
                }
            AddGap(difference.Left, startLeft, left.Length);
            AddGap(difference.Right, startRight, right.Length);
            return difference;
        }

        static void AddGap(List<ComparisonTextSpan> spans, int start, int end)
        {
            if (end > start) spans.Add(new ComparisonTextSpan(start, end - start));
        }

        static int[] Symbols(string text, Dictionary<string, int> words, ref int nextSymbol)
        {
            var symbols = new int[text.Length];
            for (int index = 0; index < text.Length;)
            {
                if (!WordCharacter(text[index]))
                {
                    symbols[index] = text[index];
                    index++;
                    continue;
                }
                int start = index++;
                while (index < text.Length && WordCharacter(text[index])) index++;
                string word = text.Substring(start, index - start);
                int symbol;
                if (!words.TryGetValue(word, out symbol))
                {
                    symbol = nextSymbol;
                    nextSymbol += word.Length;
                    words.Add(word, symbol);
                }
                // A letter or digit matches only at the same offset of the same complete word.
                for (int offset = 0; offset < word.Length; offset++) symbols[start + offset] = symbol + offset;
            }
            return symbols;
        }

        static bool WordCharacter(char character) { return Char.IsLetterOrDigit(character) || character == '$'; }
    }
}
