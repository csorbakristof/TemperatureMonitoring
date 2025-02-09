
namespace ExcelExportMerger
{
    internal class ExternalTemperatureAndGasConsumptionReport : IReport
    {
        public string ReportName => "ExtTempGas";
        private string externalTemperatureDeviceName;
        private readonly double heatingTargetTemperature;

        public TempHumValue[] Measurements { get; }
        public GasMeterValue[] GasMeterValues { get; }

        public ExternalTemperatureAndGasConsumptionReport(string externalTemperatureDeviceName, 
            double heatingTargetTemperature, TempHumValue[] measurements, GasMeterValue[] gasMeterValues)
        {
            this.externalTemperatureDeviceName = externalTemperatureDeviceName;
            this.heatingTargetTemperature = heatingTargetTemperature;
            Measurements = measurements;
            GasMeterValues = gasMeterValues;
        }

        public IEnumerable<string> GetColumnTitles() => new[] { "Day", "Mean ext temp", "Gas per day", "GasPerDeg" };

        public IEnumerable<string[]> GetReportRows()
        {
            // Note: values are averages between samplings, done by the measurement device, so no need for duration compensation
            TempHumValue[] externalTemperatureValues = Measurements.Where(m => m.DeviceName.Contains(externalTemperatureDeviceName)).ToArray();

            DateTime? previousTimestamp = null;
            double previousGasMeterValue = 0;
            foreach (var gasMeterValue in GasMeterValues)
            {
                var timestamp = gasMeterValue.Date;
                if (previousTimestamp == null)
                {
                    previousTimestamp = timestamp;
                    previousGasMeterValue = gasMeterValue.Value;
                    continue;
                }

                double? meanExternalTemperatureOffset = externalTemperatureValues.Where(m => m.Time >= previousTimestamp && m.Time < timestamp).Average(m => m.Temperature - heatingTargetTemperature);

                if (!meanExternalTemperatureOffset.HasValue)
                {
                    // No external temperature data for this period
                    previousGasMeterValue = gasMeterValue.Value;
                    previousTimestamp = timestamp;
                    continue;
                }

                double dayCount = (timestamp - previousTimestamp).Value.TotalDays;
                double gasDelta = gasMeterValue.Value - previousGasMeterValue;
                double gasDeltaPerDay = gasDelta / dayCount;

                double gasPerDeg = gasDeltaPerDay / meanExternalTemperatureOffset.Value;

                previousGasMeterValue = gasMeterValue.Value;
                previousTimestamp = timestamp;

                yield return new[] {
                    gasMeterValue.Date.ToString("yyyy-MM-dd"),
                    meanExternalTemperatureOffset.ToString() ?? "N/A",
                    gasDeltaPerDay.ToString(),
                    gasPerDeg.ToString(),
                };
            }
        }
    }
}