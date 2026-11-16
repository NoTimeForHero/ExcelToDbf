using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToDbf.Utils
{
    public class DbfHelper
    {
        private static readonly Regex regDate = new Regex(@"(\d{2,4})\.(\d{2})\.(\d{2,4})", RegexOptions.Compiled);
        private static readonly Regex regNumber = new Regex(@"^-?\d+(?:[\.,]\d+)?$", RegexOptions.Compiled);

        public static bool IsNumber(string input) => regNumber.IsMatch(input);

        public static string ToDate(string input)
        {
            if (input == null) return null;
            var match = regDate.Match(input);
            if (!match.Success) return null;
            var parts = match.Groups.Cast<Group>().Skip(1).ToArray();

            var builder = new StringBuilder();

            if (parts[0].Length == 4) // YYYY.MM.DD
            {
                builder.Append(parts[0]).Append("-").Append(parts[1]).Append("-").Append(parts[2]);
                return builder.ToString();
            }

            if (parts[2].Length == 4) // DD.MM.YYYY
            {
                builder.Append(parts[2]).Append("-").Append(parts[1]).Append("-").Append(parts[0]);
                return builder.ToString();
            }

            throw new InvalidOperationException($"Invalid date format: {input}");
        }
    }
}
