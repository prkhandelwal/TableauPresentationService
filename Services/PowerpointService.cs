/*
 * Created By Pratik Khandelwal
 */
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using PresentationService.Models;
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PresentationService.Services
{
    public interface IPowerpointService
    {
        FileStreamResult GeneratePPT(List<ViewData> viewDataList);
    }

    public class PowerpointService : IPowerpointService
    {
        private readonly ILogger _logger;
        public PowerpointService(ILogger<PowerpointService> logger)
        {
            _logger = logger;
        }
        public FileStreamResult GeneratePPT(List<ViewData> viewDataList)
        {
            //Creates a PowerPoint instance
            IPresentation pptxDoc = Presentation.Create();

            ExcelEngine excelEngine = new ExcelEngine();
            IApplication excelApplication = excelEngine.Excel;
            excelApplication.DefaultVersion = ExcelVersion.Excel2016;

            foreach (var item in viewDataList)
            {
                if (String.IsNullOrEmpty(item.ChartType))
                {
                    continue;
                }
                IWorkbook wb = excelApplication.Workbooks.Open(item.DataStream, ExcelOpenType.CSV);
                IWorksheet worksheet = wb.Worksheets[0];
                worksheet.Name = item.Name;
                DataTable dataTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames | ExcelExportDataTableOptions.DetectColumnTypes);
                if (dataTable == null)
                {
                    continue;
                }
                wb.Close();
                ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
                //Handle error here to continue after chart testing is done.
                try
                {
                    if (item.ChartType.Equals("Table"))
                    {
                        CreateTable(dataTable, slide, item);
                    }
                    else
                    {
                        CreateChart(slide, dataTable, item);
                    }
                }
                catch (Exception e)
                {
                    _logger.LogError(e.Message);
                    continue;
                }
            }

            MemoryStream stream = new MemoryStream();
            pptxDoc.Save(stream);

            //Set the position as '0'.
            stream.Position = 0;

            //Download the PowerPoint file in the browser
            FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/powerpoint");
            fileStreamResult.FileDownloadName = "Editable-" + DateTime.Now.ToString() + ".pptx";
            //WorkbookTuple.Item2.Close();
            pptxDoc.Close();
            return fileStreamResult;
        }

        private static void CreateTable(DataTable dataTable, ISlide slide, ViewData viewData)
        {
            //Filtering Data Table
            List<string> removeColumns = new List<string>();
            string paName = viewData.PrimaryAxis;
            foreach (DataColumn item in dataTable.Columns)
            {
                if (!(viewData.SerieList.Contains(item.ColumnName.Replace(" ", "")) || viewData.PrimaryAxis.Contains(item.ColumnName.Replace(" ", ""))))
                {
                    removeColumns.Add(item.ColumnName);
                }
                if (viewData.PrimaryAxis.Contains(item.ColumnName.Replace(" ", "")))
                {
                    paName = item.ColumnName;
                }
            }
            foreach (var columnName in removeColumns)
            {
                dataTable.Columns.Remove(columnName);
            }


            //Add a table to the slide
            int nColumns = dataTable.Columns.Count;
            int nRows = dataTable.Rows.Count;
            ITable table = slide.Shapes.AddTable(nRows + 1, nColumns + 1, 100, 120, 300, 200);
            //Initialize index values to add text to table cells
            int rowIndex, colIndex = 1;

            //TODO: GET column titles from Primary-Axis 
            foreach (IRow rows in table.Rows)
            {
                rowIndex = 1;

                foreach (ICell cell in rows.Cells)
                {
                    cell.TextBody.AddParagraph(dataTable.Rows[rowIndex][colIndex].ToString());
                    rowIndex++;
                }
                colIndex++;
            }
        }

        private void CreateChart(ISlide slide, DataTable dataTable, ViewData viewData)
        {
            IPresentationChart chart = slide.Charts.AddChart(100, 10, 700, 500);

            //Cleaning DataTable (Let's keep it common for all kinds)
            List<string> removeColumns = new List<string>();
            string paName = viewData.PrimaryAxis;
            foreach (DataColumn item in dataTable.Columns)
            {
                if (!(viewData.SerieList.Contains(item.ColumnName) || viewData.PrimaryAxis.Contains(item.ColumnName.Replace(" ", ""))))
                {
                    removeColumns.Add(item.ColumnName);
                }
                if (viewData.PrimaryAxis.Contains(item.ColumnName.Replace(" ", "")))
                {
                    paName = item.ColumnName;
                }
            }
            foreach (var columnName in removeColumns)
            {
                dataTable.Columns.Remove(columnName);
            }

            //Adding Data to Chart
            int nColumns = dataTable.Columns.Count;
            int nRows = dataTable.Rows.Count;

            for (int i = 1; i <= nColumns; i++)
            {
                chart.ChartData.SetValue(1, i, dataTable.Columns[i - 1]);
            }
            for (int i = 1; i < nRows; i++)
            {
                for (int j = 1; j <= nColumns; j++)
                {
                    chart.ChartData.SetValue(i + 1, j, dataTable.Rows[i - 1].ItemArray[j - 1]);
                }
            }

            //Adding page elements
            //chart.ChartType = GetChartType(viewData.ChartType);
            chart.ChartTitle = dataTable.TableName;

            if (viewData.ChartType.Equals("TreeMap"))
            {
                chart.ChartType = GetChartType(viewData.ChartType);
                chart.DataRange = chart.ChartData[1, 1, nRows, nColumns];
            }
            else if (viewData.ChartType.Equals("LineAndBar"))
            {
                CreateLineAndBar(dataTable, viewData, chart, nRows);
            }
            else if (viewData.ChartType.Equals("LineAndScatter"))
            {
                CreateLineAndScatter(dataTable, viewData, chart, nRows);
            }
            else if (viewData.ChartType.Equals("BarAndScatter"))
            {
                CreateBarAndScatter(dataTable, viewData, chart, nRows);
            }
            else
            {
                AddDataToCommonCharts(dataTable, viewData, chart, paName, nRows);
            }
        }

        private void CreateBarAndScatter(DataTable dataTable, ViewData viewData, IPresentationChart chart, int nRows)
        {
            //Serie1
            IOfficeChartSerie serie1 = chart.Series.Add(viewData.SerieList[0]);
            int columnIndex = dataTable.Columns.IndexOf(viewData.SerieList[0]);
            serie1.Values = chart.ChartData[2, columnIndex, nRows, columnIndex];
            serie1.SerieType = OfficeChartType.Column_Clustered;

            //Serie2
            IOfficeChartSerie serie2 = chart.Series.Add(viewData.SerieList[1]);
            int columnIndex2 = dataTable.Columns.IndexOf(viewData.SerieList[1]);
            serie2.Values = chart.ChartData[2, columnIndex2, nRows, columnIndex2];
            serie2.SerieType = OfficeChartType.Scatter_Line_Markers;
        }

        private void CreateLineAndScatter(DataTable dataTable, ViewData viewData, IPresentationChart chart, int nRows)
        {
            //Serie1
            IOfficeChartSerie serie1 = chart.Series.Add(viewData.SerieList[0]);
            int columnIndex = dataTable.Columns.IndexOf(viewData.SerieList[0]);
            serie1.Values = chart.ChartData[2, columnIndex, nRows, columnIndex];
            serie1.SerieType = OfficeChartType.Line;

            //Serie2
            IOfficeChartSerie serie2 = chart.Series.Add(viewData.SerieList[1]);
            int columnIndex2 = dataTable.Columns.IndexOf(viewData.SerieList[1]);
            serie2.Values = chart.ChartData[2, columnIndex2, nRows, columnIndex2];
            serie2.SerieType = OfficeChartType.Scatter_Markers;
        }

        private void CreateLineAndBar(DataTable dataTable, ViewData viewData, IPresentationChart chart, int nRows)
        {
            //Serie1
            IOfficeChartSerie serie1 = chart.Series.Add(viewData.SerieList[0]);
            int columnIndex = dataTable.Columns.IndexOf(viewData.SerieList[0]);
            serie1.Values = chart.ChartData[2, columnIndex, nRows, columnIndex];
            serie1.SerieType = OfficeChartType.Column_Clustered;

            //Serie2
            IOfficeChartSerie serie2 = chart.Series.Add(viewData.SerieList[1]);
            int columnIndex2 = dataTable.Columns.IndexOf(viewData.SerieList[1]);
            serie2.Values = chart.ChartData[2, columnIndex2, nRows, columnIndex2];
            serie2.SerieType = OfficeChartType.Line;
        }

        private void AddDataToCommonCharts(DataTable dataTable, ViewData viewData, IPresentationChart chart, string paName, int nRows)
        {
            chart.ChartType = GetChartType(viewData.ChartType);
            foreach (string item in viewData.SerieList ?? Enumerable.Empty<string>())
            {
                IOfficeChartSerie serie = chart.Series.Add(item);
                int columnIndex = dataTable.Columns.IndexOf(item) + 1;
                serie.Values = chart.ChartData[2, columnIndex, nRows, columnIndex];
            }
            var paCn = dataTable.Columns.IndexOf(paName) + 1;
            chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, paCn, nRows, paCn];

            chart.Refresh();
        }

        private OfficeChartType GetChartType(string chartType)
        {
            switch (chartType)
            {
                case "LineChart":
                    return OfficeChartType.Line;
                case "ColumnChart":
                    return OfficeChartType.Column_Clustered;
                case "PieChart":
                    return OfficeChartType.Doughnut;
                case "BarChart":
                    return OfficeChartType.Bar_Clustered;
                case "AreaChart":
                    return OfficeChartType.Area;
                case "ScatterChart":
                    return OfficeChartType.Scatter_Markers;
                case "BoxAndWhiskerChart":
                    return OfficeChartType.BoxAndWhisker;
                case "WaterFall":
                    return OfficeChartType.WaterFall;
                case "TreeMap":
                    return OfficeChartType.TreeMap;
                case "LineAndScatter":
                    return OfficeChartType.Combination_Chart;
                case "LineAndBar":
                    return OfficeChartType.Combination_Chart;
                case "BarAndScatter":
                    return OfficeChartType.Combination_Chart;

                default:
                    return OfficeChartType.Line;
            }
        }

    }
}