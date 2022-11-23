using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Utils
{
    public class URLBuilder
    {
        private string root;
        private string path;
        private readonly List<string> segments = new List<string>();

        public URLBuilder Append(string part)
        {
            if (string.IsNullOrEmpty(part)) return this;
            var type = GetType(part);
            switch (type)
            {
                case SegmentType.FullUrl:
                    var uri = new Uri(part);
                    var absPath = uri.AbsolutePath == "/" ? null : uri.AbsolutePath;
                    root = absPath == null ? uri.AbsoluteUri : uri.AbsoluteUri.Replace(absPath, "");
                    path = absPath ?? path;
                    break;
                case SegmentType.Absolute:
                    path = part;
                    break;
                case SegmentType.Relative:
                    segments.Add(part);
                    break;
            }
            return this;
        }

        public string Build()
        {
            if (root == null) throw new InvalidOperationException("Не удалось найти корень URL!");
            path = path ?? "/";
            var fullPath = root.TrimEnd('/') + '/' + path.TrimStart('/');
            if (segments.Count < 1) return fullPath;
            fullPath = fullPath.TrimEnd('/') + "/" + string.Join("/", segments.Select(x => x.TrimEnd('/')));
            return fullPath;
        }

        private SegmentType GetType(string part)
        {
            if (part.Contains("://")) return SegmentType.FullUrl;
            return part.StartsWith("/") ? SegmentType.Absolute : SegmentType.Relative;
        }

        private enum SegmentType
        {
            FullUrl,
            Absolute,
            Relative,
        }
    }
}
