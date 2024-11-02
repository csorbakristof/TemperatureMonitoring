using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExportMerger
{
    internal class RawRounded5MinReport : RawReport
    {
        public int RoundedMinuteCount { get; set; } = 5;    // Time rounded to 5 minutes
        public override string ReportName => "Raw 5Min";

        public RawRounded5MinReport(TempHumValue[] measurements) : base(measurements)
        {
        }

        public override IEnumerable<string[]> GetReportRows()
        {
            var values = new Dictionary<string, string>();  // Device name - temperature

            DateTime lastTimestamp = new DateTime(0);
            for (int i = 0; i < measurements.Length; i++)
            {
                var time = RoundTime(measurements[i].Time);

                // If time changed, output the row
                if (lastTimestamp != time)
                {
                    if (values.Count() > 0)
                    {
                        yield return AssembleRow(lastTimestamp, values);
                        values.Clear();
                    }
                    lastTimestamp = time;
                }
                values[measurements[i].DeviceName] = $"{measurements[i].Temperature:F1}";
            }

            if (values.Count() > 0)
                yield return AssembleRow(lastTimestamp, values);
        }

        private DateTime RoundTime(DateTime time)
        {
            var minute = time.Minute - time.Minute % RoundedMinuteCount;
            return new DateTime(time.Year, time.Month, time.Day, time.Hour, minute, 0);
        }
    }
}
