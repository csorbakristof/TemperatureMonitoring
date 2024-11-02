using System.Diagnostics;

namespace ExcelExportMerger
{
    public class HeatingCyclesReport : IReport
    {
        private string ZoneName;
        private HeatingCycle[] cycles;

        public string ReportName => $"Heating cycles - {ZoneName}";

        public HeatingCyclesReport(string zoneName, HeatingCycle[] heatingCycles)
        {
            ZoneName = zoneName;
            cycles = heatingCycles;
        }

        public IEnumerable<string> GetColumnTitles() => new[] { "Day", "Start of cycle", "Duration of cycle", "Time until next start" };

        public IEnumerable<string[]> GetReportRows()
        {
            // Return results
            for (int i = 0; i < cycles.Length - 2; i++)
            {
                var startTimeDifferenceMinutes = (int)(cycles[i + 1].StartTime - cycles[i].StartTime).TotalMinutes;
                Debug.WriteLine($"Day {cycles[i].StartTime.Date}: Cycle duration {cycles[i].DurationMinutes} min, next start {startTimeDifferenceMinutes} minutes later");

                yield return new[] {
                    cycles[i].StartTime.Date.ToString("yyyy-MM-dd"),
                    cycles[i].StartTime.TimeOfDay.ToString(),
                    cycles[i].DurationMinutes.ToString(),
                    startTimeDifferenceMinutes.ToString()
                };
            }
        }
    }
}
