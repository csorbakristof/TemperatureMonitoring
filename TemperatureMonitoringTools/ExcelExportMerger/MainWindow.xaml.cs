using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Windows;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.InkML;

namespace ExcelExportMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private List<TempHumValue> measurements = new List<TempHumValue>();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "All files|*.*";
            dialog.Multiselect = true;
            if (dialog.ShowDialog() == true)
                foreach(var filename in dialog.FileNames)
                    LoadXlsx(filename);

            var deviceNames = measurements.Select(m => m.DeviceName).Distinct().OrderBy(n => n).ToList();

            Debug.WriteLine("Device names:");
            foreach (var name in deviceNames)
                Debug.WriteLine(name);

            var measTimeOrder = measurements.OrderBy(m => m.Timestamp).ToArray();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("All temperatures");
            worksheet.Cell(1, 1).Value = "Timestamp";
            worksheet.Cell(1, 2).Value = "Ventillation";
            for (int i=0; i<deviceNames.Count; i++)
                worksheet.Cell(1, 3 + i).Value = deviceNames[i];

            Debug.WriteLine("Collecting measurements");
            var prevTimestamp = new DateTime(1960, 1, 1);
            int rowIndex = 2;
            for (int i = 0; i < measTimeOrder.Length; i++)
            {
                var rowTimestamp = measTimeOrder[i].Timestamp;
                if (rowTimestamp - prevTimestamp < new TimeSpan(0, 5, 0))   // Not older than 5 minutes from previous
                    continue;
                prevTimestamp = rowTimestamp;

                worksheet.Cell(rowIndex, 1).Value = rowTimestamp;
                worksheet.Cell(rowIndex, 2).Value = 10.0 + GetVentillationLevel(rowTimestamp) * 5.0;

                for (int j = 0; j < deviceNames.Count; j++)
                {
                    var value = measurements.FirstOrDefault(m => m.DeviceName == deviceNames[j]
                        && Math.Abs((rowTimestamp-m.Timestamp).TotalMinutes)<5);
                    if (value != null)
                    {
                        Debug.WriteLine($"Device {value.DeviceName} timestamp {value.Timestamp} (diff: {(rowTimestamp - value.Timestamp).TotalMinutes} Min) temperature {value.Temperature}");
                        worksheet.Cell(rowIndex, 3 + j).Value = value.Temperature;
                        if (value.Timestamp != rowTimestamp)
                            worksheet.Cell(rowIndex, 3 + j).CreateComment().AddText(value.Timestamp.ToString());
                    }
                }

                rowIndex++;
            }

            var saveDialog = new SaveFileDialog();
            saveDialog.Filter = "XLSX files|*.xlsx";
            if (saveDialog.ShowDialog() == true)
            {
                workbook.SaveAs(saveDialog.FileName);
                Debug.WriteLine("Saved to " + saveDialog.FileName);
            }
        }

        private double GetVentillationLevel(DateTime rowTimestamp)
        {
            var time = rowTimestamp.TimeOfDay;
            if (time < new TimeSpan(9, 0, 0))
                return 0.0;
            if (time < new TimeSpan(21, 0, 0))
                return 2.0;
            return 0.0;
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
                        Timestamp = GetDateTimeFromCell(row.Cell(2)),
                        Temperature = row.Cell(3).IsEmpty() ? null : GetFloatFromCell(row.Cell(3)),
                        Humidity = row.Cell(4).IsEmpty() ? null : GetFloatFromCell(row.Cell(4)),
                        Comment = row.Cell(6).GetString()
                    };
                    measurements.Add(measurement);
                }
            }
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
    }
}