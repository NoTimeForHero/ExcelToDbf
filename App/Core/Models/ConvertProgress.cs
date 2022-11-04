using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core.Models
{
    internal class ConvertProgress
    {
        public int DocumentCurrent { get; set; }
        public int DocumentTotal { get; set; }
        public int FilesTotal { get; set; }
        public int FilesCurrent { get; set; }

        public string LocalText { get; set; } = "";
        public string GlobalText { get; set; } = "";

        public override string ToString() =>
            "[ConvertProgress " +
            $"Files=[{FilesCurrent}/{FilesTotal}]," +
            $"Document=[{DocumentCurrent}/{DocumentTotal}]" +
            $"Global=\"{GlobalText}\"" +
            $"Local=\"{LocalText}\"" +
            "]";

        public void GlobalInitialize(int filesTotal)
        {
            FilesCurrent = 0;
            FilesTotal = filesTotal;
        }

        public void FileInitialize(int current, string local, string global)
        {
            FilesCurrent = current;
            DocumentCurrent = 0;
            DocumentTotal = 0;
            GlobalText = global;
            LocalText = local;
        }

        public void SetProgress(int current, int max, string message)
        {
            DocumentCurrent = current;
            DocumentTotal = max;
            LocalText = message;
        }

        public void Reset()
        {
            DocumentCurrent = 0;
            DocumentTotal = 0;
            FilesTotal = 0;
            FilesCurrent = 0;
            LocalText = "";
            GlobalText = "";
        }
    }
}
