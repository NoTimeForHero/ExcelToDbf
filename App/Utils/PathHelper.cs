using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Utils
{
    internal class PathHelper
    {

        public List<string> components;

        public int Count => components.Count;

        public PathHelper(DirectoryInfo input)
        {
            components = Split(input);
        }

        public string GetLevel(int index)
        {
            int last = components.Count - 1;
            if (last < index)
                return null;
            return components[last - index];
        }

        protected List<string> Split(DirectoryInfo path)
        {
            if (path == null) throw new ArgumentNullException(nameof(path));
            var ret = new List<string>();
            if (path.Parent != null) ret.AddRange(Split(path.Parent));
            ret.Add(path.Name);
            return ret;
        }
    }
}
