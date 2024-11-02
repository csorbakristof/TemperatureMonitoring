namespace ExcelExportMerger
{
    public class HeatingCycle
    {
        public string ZoneName { get; set; }
        public DateTime StartTime { get; set; }
        public int DurationMinutes { get; set; }    // Time until start of next cycle
    }
}
