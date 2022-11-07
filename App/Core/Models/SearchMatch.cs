// File: SearchMatch.cs
// Created by NoTimeForHero, 2022
// Distributed under the Apache License 2.0

namespace ExcelToDbf.Core.Models
{
    public class SearchMatch
    {
        public int? Y { get; private set; }
        public int? X { get; private set; }
        public string Expected { get; private set; }
        public string Got { get; private set; }
        public bool Matches { get; private set; }

        private SearchMatch() {}

        public static SearchMatch Make(string expected, string got, bool matches) => new SearchMatch
        {
            Expected = expected,
            Got = got,
            Matches = matches
        };

        public override string ToString()
        {
            var position = X.HasValue && Y.HasValue ? $"Y={Y},X={X}" : string.Empty;
            return $"SearchMatch[{position}, Matches={Matches}, Expected=\"{Expected}\", Got=\"{Got}\"]";
        }

        public SearchMatch With(int y, int x)
        {
            Y = y;
            X = x;
            return this;
        }
    }
}