using System.Diagnostics;

namespace ExcelExportMerger
{
    internal class RawReport : IReport
    {
        protected TempHumValue[] measurements;
        protected string[] deviceNames;

        public virtual string ReportName => "Raw";

        public RawReport(TempHumValue[] measurements)
        {
            this.measurements = measurements;
            deviceNames = measurements.Select(m => m.DeviceName).Distinct().OrderBy(n => n).ToArray();
        }

        public IEnumerable<string> GetColumnTitles()
        {
            yield return "Timestamp";
            yield return "DateTime";
            foreach (var deviceName in deviceNames)
                yield return deviceName;
        }

        public virtual IEnumerable<string[]> GetReportRows()
        {
            var values = new Dictionary<string, string>();  // Device name - temperature

            DateTime lastTimestamp = new DateTime(0);
            long lastTimestampLong = 0;
            for (int i = 0; i < measurements.Length; i++)
            {
                var time = measurements[i].Time;

                // If time changed, output the row
                if (lastTimestamp != time)
                {
                    if (values.Count()>0)
                    {
                        yield return AssembleRow(lastTimestampLong, lastTimestamp, values);
                        values.Clear();
                    }
                    lastTimestamp = time;
                    lastTimestampLong = measurements[i].Timestamp;
                }
                if (values.ContainsKey(measurements[i].DeviceName))
                {
                    Debug.WriteLine($"Measurement {i}: device {measurements[i].DeviceName} already exists for this time: {time}");
                    Debug.WriteLine($"Old temperature is {values[measurements[i].DeviceName]}, new is {measurements[i].Temperature:F1}");
                }
                values[measurements[i].DeviceName] = $"{measurements[i].Temperature:F1}";
            }

            if (values.Count() > 0)
                yield return AssembleRow(lastTimestampLong, lastTimestamp, values);
        }

        protected string[] AssembleRow(long timestamp, DateTime time, Dictionary<string, string> values)
        {
            List<string> txts = new List<string>();
            txts.Add(timestamp.ToString());
            txts.Add(time.ToString("yyyy-MM-dd HH:mm"));
            for (int i=0; i < deviceNames.Length; i++)
                txts.Add( values.ContainsKey(deviceNames[i]) ? values[deviceNames[i]] : "");
            return txts.ToArray();
        }
    }
}