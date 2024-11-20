using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelExportMerger
{
    internal class RawDataLoader
    {
        public int KeepLastNDaysOnly { get; set; } = 10;    // May be overridden before the load...
        private List<TempHumValue> measurements = new List<TempHumValue>();

        public TempHumValue[] Load(string[] filenames)
        {
            measurements = new List<TempHumValue>();

            string extension = Path.GetExtension(filenames.First()).ToLower();
            if (extension == ".csv")
                foreach (var filename in filenames)
                    LoadCsv(filename);
            else if (extension == ".zip")
            {
                if (filenames.Length > 1)
                    MessageBox.Show("Only one ZIP file is supported");
                else
                    LoadZip(filenames.First());
            }
            else
                MessageBox.Show("Unsupported file type");

            DropOldMeasurements();

            CheckDifferenceOfLatestMeasurementsPerDevice();

            return measurements.OrderBy(m => m.Timestamp).ToArray();
        }

        private void CheckDifferenceOfLatestMeasurementsPerDevice()
        {
            // Get the latest Time for each device
            var latestMeasurements = measurements.GroupBy(m => m.DeviceName).Select(g => g.Max(m => m.Time)).ToArray();
            var minTime = latestMeasurements.Min();
            var maxTime = latestMeasurements.Max();
            if (maxTime - minTime > TimeSpan.FromMinutes(10))
                MessageBox.Show("There is more than 10 minutes difference between the latest measurements of the devices", "Last measurements with big time difference!");
        }

        #region Loading functions
        private void LoadZip(string v)
        {
            // Unpack zip into a temporary directory
            var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);
            System.IO.Compression.ZipFile.ExtractToDirectory(v, tempDir);
            // Load all CSV files from the temporary directory
            foreach (var file in Directory.EnumerateFiles(tempDir, "*.csv", SearchOption.AllDirectories))
                LoadCsv(file);
            // Delete the temporary directory
            Directory.Delete(tempDir, true);
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

        private void DropOldMeasurements()
        {
            var minDate = DateTime.Now - TimeSpan.FromDays(KeepLastNDaysOnly);
            measurements.RemoveAll(m => m.Time < minDate);
            Debug.WriteLine($"Removed all measurements older than {minDate}");
        }
    }
}
