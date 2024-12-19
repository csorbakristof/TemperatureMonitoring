using ClosedXML.Excel;

namespace ExcelExportMerger
{
    /// <summary>
    /// The ExcelExport class provides functionality to create an Excel workbook,
    /// add reports to it, and save the workbook to a file. It uses the ClosedXML
    /// library to handle Excel file operations.
    /// </summary>
    public class ExcelExport
    {
        private XLWorkbook workbook;

        public ExcelExport()
        {
            this.workbook = new XLWorkbook();
        }

        public void AddReport(IReport report)
        {
            var ws = workbook.Worksheets.Add(report.ReportName);

            int column = 0;
            foreach (var title in report.GetColumnTitles())
            {
                column++;
                ws.Cell(1, column).Value = title;
                ws.Cell(1, column).Style.Font.Bold = true;
            }

            int row = 1;
            foreach (string[] reportRow in report.GetReportRows())
            {
                row++;
                column = 0;
                foreach (var value in reportRow)
                {
                    column++;

                    if (double.TryParse(value, out double doubleValue))
                    {
                        ws.Cell(row, column).Value = doubleValue;
                    }
                    else
                    {
                        ws.Cell(row, column).Value = value;
                    }

                }
            }
        }

        public void Save(string filename)
        {
            workbook.SaveAs(filename);
        }
    }
}
