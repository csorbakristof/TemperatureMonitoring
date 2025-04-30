using System.Globalization;

namespace DataMergerTool
{
    internal class TempHumLoader
    {
        private List<TempHumValue> measurements = new List<TempHumValue>();

        public TempHumValue[] Load(string directoryName)
        {
            measurements.Clear();

            var filenames = Directory.EnumerateFiles(directoryName, "*.zip", SearchOption.AllDirectories).ToArray();


            foreach(var filename in filenames)
                LoadZip(filename);

            return measurements.OrderBy(m => m.Timestamp)
                .Distinct(new MeasurementDuplicateFiler())
                .ToArray();
        }

        class MeasurementDuplicateFiler : IEqualityComparer<TempHumValue>
        {
            public bool Equals(TempHumValue x, TempHumValue y)
            {
                return x.Timestamp == y.Timestamp && x.DeviceName == y.DeviceName;
            }

            public int GetHashCode(TempHumValue obj)
            {
                return obj.Timestamp.GetHashCode();
            }
        }

        #region Loading functions
        private void LoadZip(string filename)
        {
            Console.WriteLine($"Loading file: {filename}");
            // Unpack zip into a temporary directory
            var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);
            System.IO.Compression.ZipFile.ExtractToDirectory(filename, tempDir);
            // Load all CSV files from the temporary directory
            foreach (var file in Directory.EnumerateFiles(tempDir, "*.csv", SearchOption.AllDirectories))
                LoadCsv(file);
            // Delete the temporary directory
            Directory.Delete(tempDir, true);
            Console.WriteLine($"Loaded. Total measurement count: {measurements.Count}");
        }

        private void LoadCsv(string filename)
        {
            // Open CSV file and read all lines
            var lines = File.ReadAllLines(filename);
            string deviceName = lines[0];
            // lines[1] is header
            for (int i = 2; i < lines.Length; i++)
            {
                var values = lines[i].Split(';');
                long timestamp = long.Parse(values[0]);
                var measurement = new TempHumValue
                {
                    DeviceName = deviceName,
                    Timestamp = timestamp,
                    Time = GetDateTimeFromCsvTimestamp(timestamp),
                    Temperature = float.Parse(values[1], CultureInfo.InvariantCulture),
                    Humidity = float.Parse(values[2], CultureInfo.InvariantCulture),
                };
                measurements.Add(measurement);
            }
        }

        private DateTime GetDateTimeFromCsvTimestamp(long csvSecondsTimestamp)
        {
            // Convert seconds to ticks
            long timestamp = csvSecondsTimestamp * TimeSpan.TicksPerSecond;

            // Unix epoch in ticks
            long unixEpochTicks = 621355968000000000;

            // Create DateTime instance
            var dateTime = new DateTime(unixEpochTicks + timestamp, DateTimeKind.Utc);

            return dateTime;
        }
        #endregion
    }
}
