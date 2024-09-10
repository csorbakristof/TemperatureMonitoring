using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExportMerger
{
    internal class TempHumValue
    {
        public string DeviceName { get; set; }
        public DateTime Timestamp { get; set; }
        public float? Temperature { get; set; }
        public float? Humidity { get; set; }
        public string Comment { get; set; }
    }
}
