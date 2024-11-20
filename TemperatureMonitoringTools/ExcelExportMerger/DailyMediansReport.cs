
using System.Text;

namespace ExcelExportMerger
{
    internal class DailyMediansReport : RawReport
    {

        public override string ReportName => "Daily medians";

        public DailyMediansReport(TempHumValue[] measurements) : base(measurements)
        {
        }

        public override IEnumerable<string[]> GetReportRows()
        {
            var DayDevice2TemperatureList = new Dictionary<(DateOnly,string), List<TempHumValue>>();

            foreach (var m in measurements)
            {
                var date = new DateOnly(m.Time.Year, m.Time.Month, m.Time.Day);
                var key = (date, m.DeviceName);
                if (!DayDevice2TemperatureList.ContainsKey(key))
                    DayDevice2TemperatureList[key] = new List<TempHumValue>();
                DayDevice2TemperatureList[key].Add(m);
            }

            var minDate = DayDevice2TemperatureList.Keys.Min(k => k.Item1);
            var maxDate = DayDevice2TemperatureList.Keys.Max(k => k.Item1);

            for (var date = minDate; date <= maxDate; date = date.AddDays(1))
            {
                var values = new List<string>();
                values.Add("-");    // No timestamp used here
                values.Add(date.ToString("yyyy-MM-dd"));

                foreach (var deviceName in deviceNames)
                {
                    var key = (date, deviceName);
                    if (DayDevice2TemperatureList.ContainsKey(key))
                    {
                        var valuesForDevice = DayDevice2TemperatureList[key].Select(m => m.Temperature.Value).OrderBy(t => t).ToArray();
                        var median = valuesForDevice.Length % 2 == 0
                            ? (valuesForDevice[valuesForDevice.Length / 2 - 1] + valuesForDevice[valuesForDevice.Length / 2]) / 2
                            : valuesForDevice[valuesForDevice.Length / 2];
                        values.Add($"{median:F1}");
                    }
                    else
                    {
                        values.Add(" ");
                    }
                }
                yield return values.ToArray();
            }
        }
    }
}