using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ExcelToDbf.Utils
{
    internal class UniversalFolderSelector
    {

        public static string ShowDialog(string previousPath = null)
        {
            bool selected;
            string filename = null;

            if (!CommonFileDialog.IsPlatformSupported)
            {
                var dialog = new FolderBrowserDialog { SelectedPath = previousPath };
                DialogResult result = dialog.ShowDialog();
                selected = result == DialogResult.OK;
                if (selected) filename = dialog.SelectedPath;
            }
            else
            {
                var fullPath = previousPath != null ? Path.GetFullPath(previousPath) : null;
                var dialog = new CommonOpenFileDialog
                {
                    InitialDirectory = fullPath,
                    IsFolderPicker = true
                };
                CommonFileDialogResult result = dialog.ShowDialog();
                selected = result == CommonFileDialogResult.Ok;
                if (selected) filename = dialog.FileName;
            }

            if (!selected) return null;
            return filename;
        }

    }
}
