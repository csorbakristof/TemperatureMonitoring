namespace ExcelExportMerger
{
    internal class HeatingCycleDetector
    {
        const int MinimalMeasuredTemperatureForActiveHeating = 28;

        public static IEnumerable<HeatingCycle> DetectHeatingCycles(string zoneName, TempHumValue[] measurements)
        {
            var heatingMeasurements = measurements.Where(m => m.DeviceName.Contains(zoneName)).ToList();
            HeatingCycle? currentCycle = null;

            for (int i = 1; i < heatingMeasurements.Count - 1; i++)
            {
                var currentMeasurement = heatingMeasurements[i];
                var currentTemperature = heatingMeasurements[i].Temperature;
                var prevTemperature = heatingMeasurements[i-1].Temperature;

                bool isOn = currentTemperature > MinimalMeasuredTemperatureForActiveHeating;
                bool isTemperatureIncreasing = currentTemperature > prevTemperature;
                if (isOn && isTemperatureIncreasing && currentCycle == null)
                {
                    currentCycle = new HeatingCycle { StartTime = currentMeasurement.Time };
                }
                if (!isTemperatureIncreasing && currentCycle != null)
                {
                    currentCycle.DurationMinutes = (int)(currentMeasurement.Time - currentCycle.StartTime).TotalMinutes;
                    yield return currentCycle;
                    currentCycle = null;
                }
            }
        }
    }
}
