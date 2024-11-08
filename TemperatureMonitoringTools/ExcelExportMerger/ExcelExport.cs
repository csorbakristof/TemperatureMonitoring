﻿using ClosedXML.Excel;

namespace ExcelExportMerger
{
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
            }

            int row = 1;
            foreach (string[] reportRow in report.GetReportRows())
            {
                row++;
                column = 0;
                foreach (var value in reportRow)
                {
                    column++;
                    ws.Cell(row, column).Value = value;
                }
            }
        }

        public void Save(string filename)
        {
            workbook.SaveAs(filename);
        }
    }
}
