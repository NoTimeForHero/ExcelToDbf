namespace ExcelToDbf.Processing
{
    public sealed class DataWarning
    {
        public string Code { get; set; }
        public string Message { get; set; }
        public int SourceRow { get; set; }
        public string Field { get; set; }
        public object Actual { get; set; }
    }
}
