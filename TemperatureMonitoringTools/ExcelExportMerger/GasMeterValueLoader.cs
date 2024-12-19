
using ClosedXML.Excel;

namespace ExcelExportMerger
{
    internal class GasMeterValueLoader
    {
        internal static IEnumerable<GasMeterValue> Load(string fileName)
        {
            var workbook = new XLWorkbook(fileName);
            var ws = workbook.Worksheets.Worksheet(1);
            // Skip the first row, as it contains the column titles
            // Load the remaining rows into GasMeterValue objects
            for (int row = 2; row <= ws.RowsUsed().Count(); row++)
            {
                // First column contains the date in format YYYY. MM. DD, second column contains the time in format HHMM.
                // Create a DateTime object from these two columns.
                DateTime date = ws.Cell(row, 1).GetDateTime();
                DateTime time = ws.Cell(row, 2).GetDateTime();

                //int year = int.Parse(dateStr.Substring(0, 4));
                //int month = int.Parse(dateStr.Substring(6, 2));
                //int day = int.Parse(dateStr.Substring(10, 2));

                //int hour = int.Parse(timeStr.Substring(0, 2));
                //int minute = int.Parse(timeStr.Substring(2, 2));

                DateTime datetime = new DateTime(date.Year, date.Month, date.Day,
                    time.Hour, time.Minute, time.Second);

                yield return new GasMeterValue
                {
                    Date =  datetime,
                    Value = ws.Cell(row, 3).GetDouble()
                };
            }
        }
    }
}