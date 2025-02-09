using Microsoft.Win32;
using System.Diagnostics;
using System.Windows;

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

        private TempHumValue[] measurements = null;
        private GasMeterValue[] gasMeterValues = null;

        private void LoadThermometerLogs_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "All files|*.*";
            dialog.Multiselect = true;
            if (dialog.ShowDialog() == true)
            {
                var loader = new RawDataLoader() { KeepLastNDaysOnly = 60 };
                measurements = loader.Load(dialog.FileNames).ToArray();

                if (measurements.Length == 0)
                {
                    MessageBox.Show("No measurements loaded");
                    return;
                }
            }
        }

        private void LoadGasMeterValues_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel files|*.*";
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == true)
            {
                var loader = new RawDataLoader() { KeepLastNDaysOnly = 30 };
                gasMeterValues = GasMeterValueLoader.Load(dialog.FileName).ToArray();
            }
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {
            var report = new ExcelExport();
            // Add all needed reports here
            report.AddReport(new RawReport(measurements));
            report.AddReport(new RawRounded5MinReport(measurements));
            report.AddReport(new DailyMediansReport(measurements));
            var z2HeatingCycles = HeatingCycleDetector.DetectHeatingCycles("Z2", measurements).ToArray();
            report.AddReport(new HeatingCyclesReport("Z2", z2HeatingCycles));
            report.AddReport(new HeatingCycleCountReport("Z2", z2HeatingCycles));
            report.AddReport(new DailyMeanExternalTemperature("Terasz", measurements));

            report.AddReport(new CumulativeValueComparisons("Terasz", "Z2", 22.5, measurements, gasMeterValues, new HeatingCycleDetector()));
            report.AddReport(new ExternalTemperatureAndGasConsumptionReport("Terasz", 22.5, measurements, gasMeterValues));

            var saveDialog = new SaveFileDialog();
            saveDialog.Filter = "XLSX files|*.xlsx";
            if (saveDialog.ShowDialog() == true)
            {
                report.Save(saveDialog.FileName);
                Debug.WriteLine("Saved to " + saveDialog.FileName);
            }
        }
    }
}
