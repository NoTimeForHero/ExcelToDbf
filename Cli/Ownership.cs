using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace ExcelToDbf.Cli
{
    internal static class Ownership
    {
        public static void Register(string session, int id)
        {
            using (var process = Process.GetProcessById(id))
                File.AppendAllText(Path.Combine(session, "owners.txt"), id + "|" +
                    process.StartTime.ToUniversalTime().Ticks + "|" + process.ProcessName + Environment.NewLine);
        }

        public static void StopOwnedProcesses(string session)
        {
            var path = Path.Combine(session, "owners.txt");
            if (!File.Exists(path)) return;
            foreach (var line in File.ReadAllLines(path))
            {
                var parts = line.Split('|');
                if (parts.Length != 3 || !int.TryParse(parts[0], out var id) || !long.TryParse(parts[1], out var ticks)) continue;
                try
                {
                    using (var process = Process.GetProcessById(id))
                        if (process.StartTime.ToUniversalTime().Ticks == ticks && process.ProcessName == parts[2])
                        { process.Kill(); process.WaitForExit(5000); }
                }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
            }
        }
    }
}
