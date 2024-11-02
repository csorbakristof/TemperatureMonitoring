
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExportMerger
{
    internal class HeatingCycleCountReport : IReport
    {

        private string ZoneName;
        private HeatingCycle[] cycles;

        public string ReportName => $"Heating cycle count - {ZoneName}";

        public HeatingCycleCountReport(string zoneName, HeatingCycle[] heatingCycles)
        {
            ZoneName = zoneName;
            cycles = heatingCycles;
        }

        public IEnumerable<string> GetColumnTitles() => new[] { "Day", "Start count", "Time until next start" };

        public IEnumerable<string[]> GetReportRows()
        {
            var stats = cycles.GroupBy(c => c.StartTime.Date).Select(g => new { Day = g.Key, Count = g.Count(), TotalDuration = g.Sum(c => c.DurationMinutes) }).ToList();

            foreach(var stat in stats)
            {
                yield return new[] {
                    stat.Day.ToString("yyyy-MM-dd"),
                    stat.Count.ToString(),
                    stat.TotalDuration.ToString()
                };
            }
        }
    }
}