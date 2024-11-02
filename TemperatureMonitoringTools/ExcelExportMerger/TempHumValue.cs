namespace ExcelExportMerger
{
    public class TempHumValue
    {
        public string DeviceName { get; set; }
        public long Timestamp { get; set; }
        public DateTime Time { get; set; }

        public float? Temperature { get; set; }
        public float? Humidity { get; set; }
        public string Comment { get; set; }
    }
}
