using System;
using System.Collections.Generic;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Core.Services.Scripts.Data;
using ExcelToDbf.Utils.Extensions;

namespace ExcelToDbf.Processing
{
    public sealed class ProcessingException : Exception
    {
        public string Code { get; }
        public string Stage { get; }
        public string Form { get; set; }
        public int? SourceRow { get; set; }
        public string Field { get; set; }
        public object Actual { get; set; }
        public ProcessingException(string code, string stage, string message, Exception inner = null) : base(message, inner)
        { Code = code; Stage = stage; }
    }

    public sealed class ConversionDetails
    {
        public string Form { get; set; }
        public int StartY { get; set; }
        public int RowsRead { get; set; }
        public int RowsSkipped { get; set; }
        public int RecordsWritten { get; set; }
        public string StopReason { get; set; } = "end-of-sheet";
        public int? StopRow { get; set; }
        public List<Dictionary<string, object>> Sample { get; } = new List<Dictionary<string, object>>();
    }

    public static class ConversionPipeline
    {
        public static ConversionDetails Write(ExcelService excel, ExcelContext context, DBFService database,
            DocForm form, string output, Action<int, int> progress = null)
        {
            var details = new ConversionDetails { Form = form.Name, StartY = form.Settings.StartY };
            if (context.TryGetContextValue<int>("startY", out var startY)) details.StartY = startY;
            if (details.StartY < 1)
                throw new ProcessingException("INVALID_START_ROW", "settings", "startY должен быть не меньше 1") { Form = form.Name };
            var stage = "dbf-header";
            int? sourceRow = null;
            try
            {
                using (var dbf = database.Make(form, output))
                {
                    dbf.WriteHeaders();
                    stage = "beforeWrite";
                    context.CallHook(form, DocForm.HookType.Before);
                    var current = details.StartY;
                    stage = "read";
                    foreach (var block in excel.IterateRanges(details.StartY, form.Settings.EndX))
                    {
                        foreach (var row in block.AsRowArray())
                        {
                            sourceRow = current++;
                            details.RowsRead++;
                            progress?.Invoke(sourceRow.Value, excel.SheetRows);
                            stage = "write";
                            Dictionary<string, object> record;
                            try { record = context.Transform(form, row); }
                            catch (StopFunctionException)
                            {
                                details.StopReason = "stop";
                                details.StopRow = sourceRow;
                                goto Finished;
                            }
                            if (record == null) details.RowsSkipped++;
                            else
                            {
                                stage = "dbf-write";
                                dbf.WriteRecord(record, sourceRow.Value);
                                if (details.Sample.Count < 5) details.Sample.Add(record);
                            }
                        }
                        stage = "read";
                    }
                Finished:
                    details.RecordsWritten = dbf.RecordsWritten;
                    sourceRow = null;
                    stage = "afterWrite";
                    context.CallHook(form, DocForm.HookType.After);
                }
                return details;
            }
            catch (ProcessingException error)
            {
                error.Form = form.Name;
                error.SourceRow = error.SourceRow ?? sourceRow;
                throw;
            }
            catch (Exception error)
            {
                throw new ProcessingException("CONVERSION_FAILED", stage, error.Message, error)
                { Form = form.Name, SourceRow = sourceRow };
            }
        }
    }
}
