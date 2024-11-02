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
        public int KeepLastNDaysOnly { get; set; } = 5;
        private List<TempHumValue> measurements = new List<TempHumValue>();

        public TempHumValue[] Load(string[] filenames)
        {
            measurements = new List<TempHumValue>();

            string extension = Path.GetExtension(filenames.First()).ToLower();
            if (extension == ".xlsx")
                foreach (var filename in filenames)
                    LoadXlsx(filename);
            else if (extension == ".csv")
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

            return measurements.OrderBy(m => m.Timestamp).ToArray();
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

        private void LoadXlsx(string filename)
        {
            using (var workbook = new XLWorkbook(filename))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();
                bool isHeader = true;
                foreach (var row in rows)
                {
                    if (isHeader)
                    {
                        isHeader = false;
                        continue;
                    }
                    if (row.Cell(1).IsEmpty() || row.Cell(2).IsEmpty())
                        continue;

                    var measurement = new TempHumValue
                    {
                        DeviceName = row.Cell(1).GetString(),
                        Time = GetDateTimeFromCell(row.Cell(2)),
                        Timestamp = row.Cell(2).GetDateTime().Ticks,
                        Temperature = row.Cell(3).IsEmpty() ? null : GetFloatFromCell(row.Cell(3)),
                        Humidity = row.Cell(4).IsEmpty() ? null : GetFloatFromCell(row.Cell(4)),
                        Comment = row.Cell(6).GetString()
                    };
                    measurements.Add(measurement);
                }
            }
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

        private float GetFloatFromCell(IXLCell cell) =>
            float.Parse(cell.GetString(), System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture);

        private readonly Regex dateTimeExtractor = new Regex(@"(\d+-\d+-\d+)T(\d+:\d+):.+");    // 2024-08-29T01:47:33.923+02:00
        private DateTime GetDateTimeFromCell(IXLCell cell)
        {
            var timeString = cell.GetString();
            var reformattedTimeString = dateTimeExtractor.Replace(timeString, @"$1 $2");
            return DateTime.Parse(reformattedTimeString);
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
