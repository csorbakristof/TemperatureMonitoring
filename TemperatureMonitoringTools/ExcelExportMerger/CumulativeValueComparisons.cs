
namespace ExcelExportMerger
{
    internal class CumulativeValueComparisons : IReport
    {
        public string ReportName => "Kumulativ";
        private string externalTemperatureDeviceName;
        private readonly string heatingTemperatureDeviceName;
        private readonly double heatingTargetTemperature;

        public TempHumValue[] Measurements { get; }
        public GasMeterValue[] GasMeterValues { get; }

        private readonly HeatingCycleDetector heatingCycleDecector;

        public CumulativeValueComparisons(string externalTemperatureDeviceName, string heatingTemperatureDeviceName,
            double heatingTargetTemperature, TempHumValue[] measurements, GasMeterValue[] gasMeterValues,
            HeatingCycleDetector heatingCycleDetector)
        {
            this.externalTemperatureDeviceName = externalTemperatureDeviceName;
            this.heatingTemperatureDeviceName = heatingTemperatureDeviceName;
            this.heatingTargetTemperature = heatingTargetTemperature;
            Measurements = measurements;
            GasMeterValues = gasMeterValues;
            this.heatingCycleDecector = heatingCycleDetector;
        }

        public IEnumerable<string> GetColumnTitles() => new[] { "Day", "Temperature", "Z2 cycle count", "Gas meter" };

        public IEnumerable<string[]> GetReportRows()
        {
            var heatingCycles = heatingCycleDecector.GetHeatingCycles(heatingTemperatureDeviceName, Measurements);

            // Note: values are averages between samplings, done by the measurement device, so no need for duration compensation
            TempHumValue[] externalTemperatureValues = Measurements.Where(m => m.DeviceName.Contains(externalTemperatureDeviceName)).ToArray();

            foreach (var gasMeterValue in GasMeterValues)
            {
                var timestamp = gasMeterValue.Date;

                // count in single zone (!)
                int heatingCycleCount = heatingCycles.Count(h => h.StartTime < timestamp);  // Assuming gas meter value is measured during no heating cycle
                double? cumulativeExternalTemperatureOffset = externalTemperatureValues.Where(m => m.Time < timestamp).Sum(m => m.Temperature - heatingTargetTemperature);
                yield return new[] {
                    gasMeterValue.Date.ToString("yyyy-MM-dd"),
                    cumulativeExternalTemperatureOffset.ToString() ?? "N/A",
                    heatingCycleCount.ToString(),
                    gasMeterValue.Value.ToString()
                };
            }
        }
    }
}