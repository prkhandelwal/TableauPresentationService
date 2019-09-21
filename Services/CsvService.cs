/*
 * Created By Pratik Khandelwal
 */
using Microsoft.AspNetCore.Mvc;
using PresentationService.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;

namespace PresentationService.Services
{
    public interface ICsvService
    {
        FileStreamResult GenerateWorkbook(List<ViewData> files);
    }

    public class CsvService : ICsvService
    {
        public FileStreamResult GenerateWorkbook(List<ViewData> files)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Excel2016;
            IWorkbook workbook = application.Workbooks.Create(1);

            foreach (var item in files)
            {
                IWorkbook wb = application.Workbooks.Open(item.DataStream, ExcelOpenType.CSV);
                IWorksheet ws = workbook.Worksheets.AddCopy(wb.Worksheets[0]);
                ws.Name = item.Name;
                wb.Close();
            }

            workbook.Worksheets[0].Remove();
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);
            workbook.Close();

            //Set the position as '0'.
            stream.Position = 0;

            FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            fileStreamResult.FileDownloadName = "CSVFile" + DateTime.Now.ToString() + ".xlsx";
            //WorkbookTuple.Item2.Close();

            return fileStreamResult;
        }
    }
}