using ConsoleApp.Models;
using System.Reflection;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp.Services
{
    internal class ExcelService
    {
        public async Task<byte[]> CreateAsync(IEnumerable<Source> sources)
        {
            var application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;
            application.ScreenUpdating = false;
            application.EnableEvents = false;
            application.ErrorCheckingOptions.BackgroundChecking = false;
            application.DisplayStatusBar = false;
            application.PrintCommunication = false;

            var workBook = (Excel._Workbook)(application.Workbooks.Add(Missing.Value));
            var count = workBook.Sheets.Count;

            var sourceWorksheet = (Excel._Worksheet)workBook.Sheets[1];
            sourceWorksheet.Name = "Source";

            var columNames = new string[]
            {
                nameof(Source.Id),
                nameof(Source.Created),
                "SpanTime",
                nameof(Source.GasTemperature),
                nameof(Source.ReferenceTemperature),
                nameof(Source.SampleTemperature),
                nameof(Source.ContainerPlatePosition),
            };

            for (int i = 0; i < columNames.Length; i++)
            {
                var columName = columNames[i];
                sourceWorksheet.Cells[1, i + 1] = columName;
            }

            var sourceCount = sources.Count();

            var cells = new object[sourceCount, columNames.Length];

            for (int i = 0; i < sourceCount; i++)
            {
                var source = sources.ElementAt(i);
                cells[i, 0] = source.Id;
                cells[i, 1] = source.Created;
                cells[i, 2] = $"={GetAlphabet(2)}{i + 2} - $B$2";
                cells[i, 3] = source.GasTemperature;
                cells[i, 4] = source.ReferenceTemperature;
                cells[i, 5] = source.SampleTemperature;
                cells[i, 6] = source.ContainerPlatePosition;
            }

            var sourceRange = sourceWorksheet.Range[sourceWorksheet.Cells[2, 1], sourceWorksheet.Cells[sourceCount + 1, columNames.Length]];
            sourceRange.Value2 = cells;

            var createdRange = sourceWorksheet.Range[sourceWorksheet.Cells[2, GetAlphabet(2)], sourceWorksheet.Cells[sourceCount + 1, GetAlphabet(2)]];
            createdRange.NumberFormat = "hh:mm:ss";

            var timeSpanRange = sourceWorksheet.Range[sourceWorksheet.Cells[2, GetAlphabet(3)], sourceWorksheet.Cells[sourceCount + 1, GetAlphabet(3)]];
            timeSpanRange.NumberFormat = "hh:mm:ss";
            
            var graphSheet = (Excel._Worksheet)workBook.Sheets.Add();
            graphSheet.Name = "Tracking";

            var chartSize = new
            {
                X0 = 10,
                X1 = 800,
                Y0 = 10,
                Y1 = 400
            };

            var chartObject = (Excel.ChartObject)graphSheet.ChartObjects().Add(
                chartSize.X0,
                chartSize.Y0,
                chartSize.X1,
                chartSize.Y1);

            var chart = chartObject.Chart;

            chart.ChartType = Excel.XlChartType.xlLine;
            chart.ChartWizard(
                Source: sourceWorksheet.Range
                [
                    sourceWorksheet.Cells[1, GetAlphabet(3)],
                    sourceWorksheet.Cells[sourceCount + 1, GetAlphabet(7)]
                ],
                Title: "Tracking");

            chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

            var seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
            var posSeries = seriesCollection.Item(4);
            posSeries.AxisGroup = Excel.XlAxisGroup.xlSecondary;

            var primaryXAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory);
            primaryXAxis.TickMarkSpacing = 300;
            primaryXAxis.TickLabelSpacing = 600;

            var primaryYAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            primaryYAxis.Crosses = Excel.XlAxisCrosses.xlAxisCrossesMaximum;
            primaryYAxis.MinimumScale = -200;
            primaryYAxis.MaximumScale = 30;

            var secondYAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary);
            secondYAxis.MinimumScale = -800;
            secondYAxis.MaximumScale = 30;

            var folder = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var fileName = "temp.xlsx";
            var filePath = $"{folder}\\{fileName}";

            // Default save locatoin is my documents
            workBook.SaveAs2(
                fileName,
                Excel.XlFileFormat.xlWorkbookDefault,
                Type.Missing,
                Type.Missing,
                false,
                false,
                Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing);

            workBook.Close(false);
            application.Quit();

            var bytes = await File.ReadAllBytesAsync(filePath);

            File.Delete(filePath);

            return bytes;
        }

        private string GetAlphabet(int number)
        {
            return ((char)(number + 64)).ToString();
        }
    }
}
