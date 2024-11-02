namespace ExcelExportMerger
{
    public interface IReport
    {
        public string ReportName { get; }
        public IEnumerable<string> GetColumnTitles();
        public IEnumerable<string[]> GetReportRows();
    }
}
