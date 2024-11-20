
using System.Diagnostics.Metrics;

namespace ExcelExportMerger
{
    internal class DailyMeanExternalTemperature : IReport
    {
        public string ReportName => "Daily external mean";

        public string ZoneName { get; }

        private TempHumValue[] measurements;

        public DailyMeanExternalTemperature(string zoneName, TempHumValue[] measurements)
        {
            this.ZoneName = zoneName;
            this.measurements = measurements;
        }

        public IEnumerable<string> GetColumnTitles() => new string[] { "Day", "Mean temperature" };

        public IEnumerable<string[]> GetReportRows()
        {
            // Find the measurements with zone name matching this.ZoneName
            var zoneMeasurements = measurements.Where(m => m.DeviceName.Contains(this.ZoneName));
            // Calculate the mean temperature for each day
            var stats = zoneMeasurements.GroupBy(m => m.Time.Date).Select(g => new { Day = g.Key, MeanTemperature = g.Average(m => m.Temperature) }).ToArray();

            foreach (var stat in stats)
            {
                yield return new string[] {
                    stat.Day.ToString("yyyy-MM-dd"),
                    stat.MeanTemperature.ToString() ?? "N/A"
                };
            }
        }
    }
}